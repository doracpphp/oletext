package oletext

import "testing"

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
