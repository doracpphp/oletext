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
	binary.LittleEndian.PutUint16(data[26:], 3)          // major version 3
	binary.LittleEndian.PutUint16(data[30:], 9)          // 512-byte sectors
	binary.LittleEndian.PutUint32(data[44:], 0xFFFFFFFA) // numFATSects (huge)
	binary.LittleEndian.PutUint32(data[48:], 0)          // firstDirSect
	binary.LittleEndian.PutUint32(data[68:], 0)          // firstDIFATSect -> sector 0
	binary.LittleEndian.PutUint32(data[72:], 0xFFFFFFFA) // numDIFATSects (huge)
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
