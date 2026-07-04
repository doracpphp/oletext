package oletext

import (
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
