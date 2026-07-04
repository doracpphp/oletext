package oletext

import "testing"

// Legacy Word (.doc) tests against the LibreOffice-generated samples
// (testdata/gen.py shapes|big).

// TestDocBodyAndShape covers body paragraph text and text-frame (shape) text.
func TestDocBodyAndShape(t *testing.T) {
	extractFileWant(t, "testdata/shapes.doc",
		"Body paragraph 本文です。",
		"Shape in Word 図形の中のテキスト",
	)
}

// TestDocLarge checks every paragraph of the large big.doc is extracted (no
// gaps, duplicates or truncation), including a 30k-char run.
func TestDocLarge(t *testing.T) {
	verifyBig(t, "testdata/big.doc", 8000, `日本語\d{6}`, []string{"LONGSTART", "LONGEND_MK"})
}

// TestCleanDocTextFields covers the field-code stripping (0x13 begin /
// 0x14 separator / 0x15 end): only the field result survives, nesting
// works, and a field with no separator (an XE index entry, all code) does
// not swallow the text that follows it.
func TestCleanDocTextFields(t *testing.T) {
	for _, tc := range []struct{ in, want string }{
		{"X\x13DATE\x142026\x15Y", "X2026Y"},
		{"A\x13IF \x13PAGE\x142\x15 > 1\x14Result\x15B", "AResultB"},
		{"P\x13XE bookmark\x15Q", "PQ"},                    // no separator: all code
		{"P\x13XE a\x15Q\x13DATE\x142026\x15R", "PQ2026R"}, // and later fields still work
	} {
		if got := cleanDocText(tc.in); got != tc.want {
			t.Errorf("cleanDocText(%q) = %q, want %q", tc.in, got, tc.want)
		}
	}
}
