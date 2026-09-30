package oletext

import (
	"encoding/binary"
	"testing"
	"time"
)

// TestParseCFBBoundedDIFAT guards against the DIFAT/FAT walk doing an
// effectively unbounded amount of work when the header counts are corrupt.
// The crafted file claims a huge numDIFATSects and points the DIFAT chain at
// a sector that loops back on itself; parseCFB must clamp the walk to the
// number of sectors the file can hold and return promptly instead of
// building a giant FAT (found by fuzzing FuzzExtract).
func TestParseCFBBoundedDIFAT(t *testing.T) {
	const sectorSize = 512
	data := make([]byte, sectorSize*5) // header + 4 sectors -> maxSects == 5
	copy(data, cfbSignature)
	binary.LittleEndian.PutUint16(data[26:], 3)             // major version 3
	binary.LittleEndian.PutUint16(data[30:], 9)             // 512-byte sectors
	binary.LittleEndian.PutUint32(data[44:], 0xFFFFFFFA)    // numFATSects (huge)
	binary.LittleEndian.PutUint32(data[48:], 0)             // firstDirSect
	binary.LittleEndian.PutUint32(data[68:], 0)             // firstDIFATSect -> sector 0
	binary.LittleEndian.PutUint32(data[72:], 0xFFFFFFFA)    // numDIFATSects (huge)
	binary.LittleEndian.PutUint32(data[2*sectorSize-4:], 0) // DIFAT chain self-loop

	done := make(chan struct{})
	go func() {
		parseCFB(data) // may return an error; it must not hang or panic.
		close(done)
	}()
	select {
	case <-done:
	case <-time.After(2 * time.Second):
		t.Fatal("parseCFB did not return within 2s on a corrupt DIFAT count")
	}
}

// TestStreamLookupScopedToRoot checks the document type is decided by the
// streams of the root storage. An embedded OLE object keeps its own streams
// in a sub-storage under the same well-known names, so a workbook with an
// embedded Word document must still be read as a workbook.
func TestStreamLookupScopedToRoot(t *testing.T) {
	data := buildCFB(map[string][]byte{
		"MBD0001/WordDocument": []byte("not a real FIB"),
		"Workbook":             minimalWorkbook("Sheet text"),
	})
	f, err := parseCFB(data)
	if err != nil {
		t.Fatalf("parseCFB: %v", err)
	}
	if f.hasStream("WordDocument") {
		t.Error("hasStream found the embedded object's WordDocument stream in the root storage")
	}
	if !f.hasStream("workbook") {
		t.Error("hasStream should match the root Workbook stream case-insensitively")
	}
	extractWant(t, data, "=== Sheet: S ===", "Sheet text")
}

// TestReadChainLoop checks a cyclic FAT chain is reported as an error rather
// than followed indefinitely.
func TestReadChainLoop(t *testing.T) {
	data := buildCFB(map[string][]byte{"A": make([]byte, 600)})
	// Layout: sector 0 = FAT, 1 = directory, 2-3 = stream A.
	binary.LittleEndian.PutUint32(data[512+1*4:], 1) // FAT[1] -> 1: the directory chain loops
	if _, err := parseCFB(data); err == nil {
		t.Error("expected an error for a looping directory chain, got nil")
	}
}
