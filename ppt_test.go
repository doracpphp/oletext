package oletext

import "testing"

// Legacy PowerPoint (.ppt) tests against the LibreOffice-generated samples
// (testdata/gen.py shapes|big).

// TestPptBodyAndShape covers placeholder (title) text and a free text-box
// (shape).
func TestPptBodyAndShape(t *testing.T) {
	extractFileWant(t, "testdata/shapes.ppt",
		"Slide Title タイトル",
		"Shape in PowerPoint 図形テキスト",
	)
}

// TestPptLarge checks every slide's text in the large big.ppt is extracted.
func TestPptLarge(t *testing.T) {
	verifyBig(t, "testdata/big.ppt", 250, `スライド\d{6}`, nil)
}
