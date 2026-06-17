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
