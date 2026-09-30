package oletext

import (
	"bytes"
	"encoding/binary"
	"strings"
	"testing"
)

// Legacy Excel (.xls) tests against the LibreOffice-generated samples
// (testdata/gen.py shapes|big).

// TestXlsBodyAndShape covers ordinary cell text and text-box (shape) text.
func TestXlsBodyAndShape(t *testing.T) {
	extractFileWant(t, "testdata/shapes.xls",
		"CellText",
		"Shape in Excel 図形テキスト",
	)
}

// TestXlsLarge checks every cell of the large big.xls is extracted, including
// a string long enough to span Continue records.
func TestXlsLarge(t *testing.T) {
	verifyBig(t, "testdata/big.xls", 30000, `和\d{6}`, []string{"LONGSTART", "LONGEND_MK"})
}

// biffRec appends one BIFF record (type, size, payload) to b.
func biffRec(b *bytes.Buffer, typ uint16, data []byte) {
	binary.Write(b, binary.LittleEndian, typ)
	binary.Write(b, binary.LittleEndian, uint16(len(data)))
	b.Write(data)
}

// minimalWorkbook builds a BIFF8 Workbook stream with one sheet "S" whose
// cell A1 holds text, followed by any extra sheet records.
func minimalWorkbook(text string, extra ...func(*bytes.Buffer)) []byte {
	var b bytes.Buffer
	biffRec(&b, recBOF, []byte{0x00, 0x06, 0x05, 0x00}) // BIFF8, workbook globals
	biffRec(&b, recBoundSheet, []byte{0, 0, 0, 0, 0, 0, 1, 0, 'S'})
	biffRec(&b, recEOF, nil)
	biffRec(&b, recBOF, []byte{0x00, 0x06, 0x10, 0x00})      // BIFF8, worksheet
	label := []byte{0, 0, 0, 0, 0, 0, byte(len(text)), 0, 0} // row, col, ixfe, cch, flags
	biffRec(&b, recLabel, append(label, text...))
	for _, fn := range extra {
		fn(&b)
	}
	biffRec(&b, recEOF, nil)
	return b.Bytes()
}

// TestXlsBoolAndError checks boolean and error cells are emitted, both as
// BoolErr records and as cached formula results.
func TestXlsBoolAndError(t *testing.T) {
	wb := minimalWorkbook("Label", func(b *bytes.Buffer) {
		biffRec(b, recBoolErr, []byte{1, 0, 0, 0, 0, 0, 1, 0})    // row 1: TRUE
		biffRec(b, recBoolErr, []byte{1, 0, 1, 0, 0, 0, 0x07, 1}) // row 1: #DIV/0!
		formula := make([]byte, 20)
		formula[0] = 2                        // row 2
		formula[6], formula[8] = 0x01, 0      // boolean FALSE
		formula[12], formula[13] = 0xFF, 0xFF // non-numeric result
		biffRec(b, recFormula, formula)
		formula[2] = 1                      // next column
		formula[6], formula[8] = 0x02, 0x2A // error #N/A
		biffRec(b, recFormula, formula)
	})
	extractWant(t, buildCFB(map[string][]byte{"Workbook": wb}),
		"Label\nTRUE\t#DIV/0!\nFALSE\t#N/A\n")
}

// hlinkPayload builds an HLink record body: ref8 + hlinkClsid +
// streamVersion + flags, followed by the flag-dependent fields in rest.
func hlinkPayload(flags uint32, rest []byte) []byte {
	d := make([]byte, 8+16+4+4)
	binary.LittleEndian.PutUint32(d[28:], flags)
	return append(d, rest...)
}

// TestHLinkURLMoniker checks the URL of a web hyperlink is emitted.
func TestHLinkURLMoniker(t *testing.T) {
	url := "http://example.com"
	body := urlMonikerCLSID[:]
	u16 := vbaUTF16(url + "\x00")
	var nb [4]byte
	binary.LittleEndian.PutUint32(nb[:], uint32(len(u16)))
	body = append(body, nb[:]...)
	body = append(body, u16...)

	x := &xlsExtractor{curRow: -1, sheetIdx: -1, atLineStart: true}
	x.onHLink(hlinkPayload(hlinkHasMoniker, body))
	if got := x.out.String(); !strings.Contains(got, url) {
		t.Errorf("URL moniker: got %q, want it to contain %q", got, url)
	}
}

// TestHLinkFileMonikerNoGarbage checks that an unparsed (non-URL) moniker
// does not make the location-string reader decode moniker internals as text.
func TestHLinkFileMonikerNoGarbage(t *testing.T) {
	body := make([]byte, 16) // an all-zero CLSID: not the URL moniker
	// FileMoniker-ish innards that would misparse as a plausible
	// HyperlinkString (length 4, then 8 bytes of UTF-16 junk).
	var n [4]byte
	binary.LittleEndian.PutUint32(n[:], 4)
	body = append(body, n[:]...)
	body = append(body, []byte{0x41, 0x00, 0x42, 0x00, 0x43, 0x00, 0x44, 0x00}...)

	x := &xlsExtractor{curRow: -1, sheetIdx: -1, atLineStart: true}
	x.onHLink(hlinkPayload(hlinkHasMoniker|hlinkHasLocationStr, body))
	if got := x.out.String(); got != "" {
		t.Errorf("file moniker: got %q, want no output", got)
	}
}
