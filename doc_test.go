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

// TestCleanDocTextTable covers table text: cells are tab-separated, a row
// mark ends the line without a trailing tab, and empty cells keep their
// column position.
func TestCleanDocTextTable(t *testing.T) {
	const row = string(rune(docRowMark))
	for _, tc := range []struct{ in, want string }{
		{"a\x07b\x07" + row + "c\x07d\x07" + row + "after\r", "a\tb\nc\td\nafter\n"},
		{"a\x07\x07c\x07" + row, "a\t\tc\n"}, // empty middle cell
		{"\x07b\x07" + row, "\tb\n"},         // empty first cell
		{"a\x07b\x07\x07", "a\tb\t\t"},       // no row information: tabs only
		{"x\r" + "a\x07" + row + "y\r", "x\na\ny\n"},
	} {
		if got := cleanDocText(tc.in); got != tc.want {
			t.Errorf("cleanDocText(%q) = %q, want %q", tc.in, got, tc.want)
		}
	}
}

// TestGrpprlHasTtp covers the Prl walk that finds sprmPFTtp, including the
// operand sizes it has to step over and a grpprl moved to the Data stream.
func TestGrpprlHasTtp(t *testing.T) {
	inTable := []byte{0x16, 0x24, 1}           // sprmPFInTable = 1
	ttp := []byte{0x17, 0x24, 1}               // sprmPFTtp = 1
	defTable := []byte{0x08, 0xD6, 3, 0, 9, 9} // sprmTDefTable, cb=3: 2 bytes follow
	huge := []byte{0x46, 0x66, 2, 0, 0, 0}     // sprmPHugePapx -> Data offset 2
	data := func() []byte { return append([]byte{0xEE, 0xEE, 3, 0}, ttp...) }
	cat := func(parts ...[]byte) []byte {
		var b []byte
		for _, p := range parts {
			b = append(b, p...)
		}
		return b
	}
	for _, tc := range []struct {
		name string
		g    []byte
		want bool
	}{
		{"row mark", cat(inTable, ttp), true},
		{"cell paragraph", inTable, false},
		{"after a long operand", cat(defTable, ttp), true},
		{"explicitly cleared", cat(ttp, []byte{0x17, 0x24, 0}), false},
		{"in the Data stream", huge, true},
		{"truncated", []byte{0x17, 0x24}, false},
	} {
		if got := grpprlHasTtp(tc.g, data); got != tc.want {
			t.Errorf("%s: grpprlHasTtp = %v, want %v", tc.name, got, tc.want)
		}
	}
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
