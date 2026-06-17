package oletext

import (
	"archive/zip"
	"bytes"
	"encoding/binary"
	"fmt"
	"os"
	"path/filepath"
	"regexp"
	"sort"
	"strings"
	"testing"
	"time"
	"unicode/utf16"
)

// This file holds the helpers shared by the (deliberately fine-grained)
// test files: each feature gets its own *_test.go, and they all draw on the
// builders and assertions below.

// ---- ZIP / OOXML package building ----

// zipBytes packs name->content parts into an in-memory ZIP, the container an
// OOXML document uses.
func zipBytes(parts map[string]string) []byte {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	for name, content := range parts {
		w, _ := zw.Create(name)
		w.Write([]byte(content))
	}
	zw.Close()
	return buf.Bytes()
}

// ---- assertions ----

// wantContains fails the test for any substring missing from got.
func wantContains(t *testing.T, got string, wants ...string) {
	t.Helper()
	for _, w := range wants {
		if !strings.Contains(got, w) {
			t.Errorf("output missing %q\n--- got ---\n%s", w, got)
		}
	}
}

// wantAbsent fails the test for any substring that leaked into got.
func wantAbsent(t *testing.T, got string, bads ...string) {
	t.Helper()
	for _, b := range bads {
		if strings.Contains(got, b) {
			t.Errorf("output should not contain %q\n--- got ---\n%s", b, got)
		}
	}
}

// extractWant runs Extract and asserts every want is present, returning the
// extracted text for any further checks (ordering, absence).
func extractWant(t *testing.T, data []byte, wants ...string) string {
	t.Helper()
	got, err := Extract(data)
	if err != nil {
		t.Fatalf("Extract: %v", err)
	}
	wantContains(t, got, wants...)
	return got
}

// extractFileWant runs ExtractFile against a testdata sample, skipping when
// the (git-ignored) file is absent.
func extractFileWant(t *testing.T, path string, wants ...string) string {
	t.Helper()
	if _, err := os.Stat(path); err != nil {
		t.Skipf("%s not present (run testdata/gen.py to create it)", path)
	}
	got, err := ExtractFile(path)
	if err != nil {
		t.Fatalf("ExtractFile(%s): %v", path, err)
	}
	wantContains(t, got, wants...)
	return got
}

// ---- VBA project building ([MS-OVBA]) ----

// compressVBALiteral produces a valid [MS-OVBA] 2.4.1 CompressedContainer
// storing data with literal tokens only (no copy tokens). Each 2048 source
// bytes form one compressed chunk.
func compressVBALiteral(data []byte) []byte {
	out := []byte{0x01} // SignatureByte
	for start := 0; start < len(data); start += 2048 {
		chunk := data[start:min(start+2048, len(data))]
		var body []byte
		for i := 0; i < len(chunk); i += 8 {
			body = append(body, 0x00) // FlagByte: the next 8 tokens are literals
			body = append(body, chunk[i:min(i+8, len(chunk))]...)
		}
		// Header: bit15 compressed, bits12-14 = 0b011, bits0-11 = dataLen-1.
		header := uint16(0x8000) | uint16(0x3000) | uint16(len(body)-1)
		out = append(out, byte(header), byte(header>>8))
		out = append(out, body...)
	}
	return out
}

// vbaRec appends one dir-stream record (Id, Size, Data).
func vbaRec(b *bytes.Buffer, id uint16, data []byte) {
	binary.Write(b, binary.LittleEndian, id)
	binary.Write(b, binary.LittleEndian, uint32(len(data)))
	b.Write(data)
}

// vbaUTF16 encodes s as little-endian UTF-16 bytes.
func vbaUTF16(s string) []byte {
	var b bytes.Buffer
	for _, r := range s {
		binary.Write(&b, binary.LittleEndian, uint16(r))
	}
	return b.Bytes()
}

// buildVBADir builds a decompressed dir stream describing a single module in
// the record layout parseVBADir expects.
func buildVBADir(streamName string, offset int) []byte {
	var b bytes.Buffer
	// PROJECTMODULES: Id(0x000F), Size=2, Count=1.
	binary.Write(&b, binary.LittleEndian, uint16(0x000F))
	binary.Write(&b, binary.LittleEndian, uint32(2))
	binary.Write(&b, binary.LittleEndian, uint16(1))

	vbaRec(&b, 0x0019, []byte(streamName)) // MODULENAME
	vbaRec(&b, 0x001A, []byte(streamName)) // MODULESTREAMNAME
	off := make([]byte, 4)
	binary.LittleEndian.PutUint32(off, uint32(offset))
	vbaRec(&b, 0x0031, off) // MODULEOFFSET

	// MODULE terminator: Id(0x002B), Size=0, plus 4 reserved bytes.
	binary.Write(&b, binary.LittleEndian, uint16(0x002B))
	binary.Write(&b, binary.LittleEndian, uint32(0))
	b.Write([]byte{0, 0, 0, 0})

	// dir terminator: Id(0x0010), Size=0.
	binary.Write(&b, binary.LittleEndian, uint16(0x0010))
	binary.Write(&b, binary.LittleEndian, uint32(0))
	return b.Bytes()
}

// ---- minimal CFB writer ----

// buildCFB writes a minimal [MS-CFB] version-3 compound file holding the
// given named streams (plus the required Root Entry). MiniCutoff is set to 0
// so every stream lives in the regular FAT, avoiding the mini stream and
// letting readChain return each stream at its exact size.
func buildCFB(streams map[string][]byte) []byte {
	const (
		ss         = 512
		endOfChain = 0xFFFFFFFE
		freeSect   = 0xFFFFFFFF
		fatSect    = 0xFFFFFFFD
	)
	names := make([]string, 0, len(streams))
	for n := range streams {
		names = append(names, n)
	}
	sort.Strings(names)

	type entry struct {
		name  string
		data  []byte
		start uint32
		nsect int
	}
	ents := make([]entry, len(names))
	next := 2 // sector 0 = FAT, sector 1 = directory
	for i, n := range names {
		d := streams[n]
		ns := (len(d) + ss - 1) / ss
		if ns == 0 {
			ns = 1
		}
		ents[i] = entry{n, d, uint32(next), ns}
		next += ns
	}
	total := next

	fat := make([]uint32, 128)
	for i := range fat {
		fat[i] = freeSect
	}
	fat[0] = fatSect
	fat[1] = endOfChain // directory: one sector
	for _, e := range ents {
		for k := 0; k < e.nsect; k++ {
			s := int(e.start) + k
			if k == e.nsect-1 {
				fat[s] = endOfChain
			} else {
				fat[s] = uint32(s + 1)
			}
		}
	}

	le := binary.LittleEndian
	buf := make([]byte, ss*(total+1)) // +1 for the 512-byte header sector
	copy(buf, cfbSignature)
	le.PutUint16(buf[26:], 3) // major version
	le.PutUint16(buf[30:], 9) // sector shift -> 512
	le.PutUint16(buf[32:], 6) // mini sector shift -> 64
	le.PutUint32(buf[44:], 1) // numFATSects
	le.PutUint32(buf[48:], 1) // firstDirSect
	le.PutUint32(buf[56:], 0) // miniCutoff = 0 -> everything regular
	le.PutUint32(buf[60:], endOfChain)
	le.PutUint32(buf[64:], 0) // numMiniFATSects
	le.PutUint32(buf[68:], endOfChain)
	le.PutUint32(buf[72:], 0) // numDIFATSects
	le.PutUint32(buf[76:], 0) // DIFAT[0]: FAT is at sector 0
	for i := 1; i < 109; i++ {
		le.PutUint32(buf[76+i*4:], freeSect)
	}
	for i, v := range fat { // FAT sector (sector 0)
		le.PutUint32(buf[ss+i*4:], v)
	}

	dirOff := ss * 2 // directory sector (sector 1)
	writeDirEntry(buf[dirOff:], "Root Entry", 5, endOfChain, 0)
	for i, e := range ents {
		writeDirEntry(buf[dirOff+(i+1)*128:], e.name, 2, e.start, uint64(len(e.data)))
	}
	for _, e := range ents {
		copy(buf[ss*(int(e.start)+1):], e.data)
	}
	return buf
}

// writeDirEntry fills a 128-byte [MS-CFB] directory entry.
func writeDirEntry(b []byte, name string, objType byte, start uint32, size uint64) {
	le := binary.LittleEndian
	u := utf16.Encode([]rune(name))
	for i, c := range u {
		le.PutUint16(b[i*2:], c)
	}
	le.PutUint16(b[64:], uint16((len(u)+1)*2)) // name length incl. terminator
	b[66] = objType
	b[67] = 1 // color = black
	le.PutUint32(b[68:], 0xFFFFFFFF)
	le.PutUint32(b[72:], 0xFFFFFFFF)
	le.PutUint32(b[76:], 0xFFFFFFFF)
	le.PutUint32(b[116:], start)
	le.PutUint64(b[120:], size)
}

// ---- large-file completeness check ----

// verifyBig extracts a large generated file and asserts every "MK%06d"
// marker 0..wantN-1 appears exactly once (no gaps, duplicates or extras),
// that the japanese marker count matches, and that any long-string boundary
// markers survive. It skips when the (git-ignored) file is absent.
func verifyBig(t *testing.T, file string, wantN int, jaRe string, longMarkers []string) {
	t.Helper()
	data, err := os.ReadFile(file)
	if err != nil {
		t.Skipf("%s not present (run testdata/gen.py to create it)", file)
	}
	st := time.Now()
	out, err := Extract(data)
	el := time.Since(st)
	if err != nil {
		t.Fatalf("%s: Extract error: %v", file, err)
	}
	t.Logf("%s: %d bytes in, %d bytes out, %v", file, len(data), len(out), el)

	mk := regexp.MustCompile(`MK(\d{6})`)
	seen := make(map[int]int)
	for _, m := range mk.FindAllStringSubmatch(out, -1) {
		var n int
		fmt.Sscanf(m[1], "%06d", &n)
		seen[n]++
	}
	missing, dup := 0, 0
	for i := 0; i < wantN; i++ {
		switch c := seen[i]; {
		case c == 0:
			if missing < 5 {
				t.Errorf("%s: missing marker MK%06d", file, i)
			}
			missing++
		case c > 1:
			dup++
		}
	}
	if missing > 0 || dup > 0 {
		t.Errorf("%s: %d missing, %d duplicated MK markers (distinct seen=%d, want=%d)", file, missing, dup, len(seen), wantN)
	}
	if extra := len(seen) - (wantN - missing); extra > 0 {
		t.Errorf("%s: %d unexpected extra MK markers", file, extra)
	}
	if jaRe != "" {
		if got := len(regexp.MustCompile(jaRe).FindAllString(out, -1)); got != wantN {
			t.Errorf("%s: japanese marker count=%d, want %d", file, got, wantN)
		}
	}
	for _, lm := range longMarkers {
		if !strings.Contains(out, lm) {
			t.Errorf("%s: long-string marker %q missing", file, lm)
		}
	}
}

// ---- fuzz seed helpers ----

// addBinarySeeds adds every Office sample under testdata as a fuzz seed.
// testdata is git-ignored, so this is a no-op on a fresh checkout; locally it
// gives the fuzzer valid files to mutate so it reaches deep code paths.
func addBinarySeeds(f *testing.F) {
	f.Helper()
	pats := []string{
		"testdata/*.doc", "testdata/*.xls", "testdata/*.ppt",
		"testdata/*.docx", "testdata/*.xlsx", "testdata/*.pptx",
	}
	for _, pat := range pats {
		matches, _ := filepath.Glob(pat)
		for _, m := range matches {
			if b, err := os.ReadFile(m); err == nil {
				f.Add(b)
			}
		}
	}
}

// streamSeeds returns the named document stream of each testdata sample, so
// the record-level fuzz targets start from real, well-formed sequences.
func streamSeeds(stream string) [][]byte {
	var out [][]byte
	for _, pat := range []string{"testdata/*.doc", "testdata/*.xls", "testdata/*.ppt"} {
		matches, _ := filepath.Glob(pat)
		for _, m := range matches {
			data, err := os.ReadFile(m)
			if err != nil {
				continue
			}
			f, err := parseCFB(data)
			if err != nil {
				continue
			}
			if s, err := f.openStream(stream); err == nil {
				out = append(out, s)
			}
		}
	}
	return out
}

// splitBytes chops b into chunks of size n (n<=0 means a single chunk),
// modelling how a BIFF string spreads across Continue records.
func splitBytes(b []byte, n int) [][]byte {
	if n <= 0 {
		return [][]byte{b}
	}
	var segs [][]byte
	for len(b) > n {
		segs = append(segs, b[:n])
		b = b[n:]
	}
	return append(segs, b)
}
