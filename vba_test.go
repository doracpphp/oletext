package oletext

import (
	"bytes"
	"testing"
)

// VBA macro tests. The decompressor cases use hand-computed [MS-OVBA] 2.4.1
// vectors; the rest build a real CFB / vbaProject.bin from the testutil
// helpers and run the full extraction path.

func TestDecompressVBACopyToken(t *testing.T) {
	// "ABC" as three literals, then a copy token (Length=3, Offset=3): "ABCABC".
	in := []byte{0x01, 0x05, 0xB0, 0x08, 'A', 'B', 'C', 0x00, 0x20}
	if got := string(decompressVBA(in)); got != "ABCABC" {
		t.Fatalf("got %q, want %q", got, "ABCABC")
	}
}

func TestDecompressVBAOverlap(t *testing.T) {
	// "A" then a copy token (Length=5, Offset=1): an overlapping run -> "AAAAAA".
	in := []byte{0x01, 0x03, 0xB0, 0x02, 'A', 0x02, 0x00}
	if got := string(decompressVBA(in)); got != "AAAAAA" {
		t.Fatalf("got %q, want %q", got, "AAAAAA")
	}
}

func TestDecompressVBAMultiChunk(t *testing.T) {
	chunk := []byte{0x05, 0xB0, 0x08, 'A', 'B', 'C', 0x00, 0x20}
	in := append([]byte{0x01}, chunk...)
	in = append(in, chunk...)
	if got := string(decompressVBA(in)); got != "ABCABCABCABC" {
		t.Fatalf("got %q, want %q", got, "ABCABCABCABC")
	}
}

func TestDecompressVBARawChunk(t *testing.T) {
	in := []byte{0x01, 0x02, 0x30, 'X', 'Y', 'Z'} // uncompressed chunk
	if got := string(decompressVBA(in)); got != "XYZ" {
		t.Fatalf("got %q, want %q", got, "XYZ")
	}
}

func TestDecompressVBARejectsBadSignature(t *testing.T) {
	if got := decompressVBA([]byte{0x00, 0x01, 0x02}); got != nil {
		t.Errorf("expected nil without the 0x01 signature, got %v", got)
	}
	if got := decompressVBA(nil); got != nil {
		t.Errorf("expected nil for empty input, got %v", got)
	}
}

func TestParseVBADir(t *testing.T) {
	var b bytes.Buffer
	vbaRec(&b, vbaIDProjectModules, []byte{0x01, 0x00})     // Count = 1
	vbaRec(&b, 0x0013, []byte{0xFF, 0xFF})                  // PROJECTCOOKIE
	vbaRec(&b, vbaIDModuleName, []byte("Macro1"))           // friendly name (MBCS)
	vbaRec(&b, vbaIDStreamName, []byte("Macro1"))           // stream name (MBCS)
	vbaRec(&b, vbaIDStreamNameUni, vbaUTF16("Macro1"))      // stream name (Unicode)
	vbaRec(&b, vbaIDModuleOffset, []byte{0x00, 0x02, 0, 0}) // TextOffset = 512
	vbaRec(&b, 0x0021, nil)                                 // MODULETYPE
	vbaRec(&b, vbaIDModuleTerm, nil)                        // terminator...
	b.Write([]byte{0, 0, 0, 0})                             // ...+ 4 reserved bytes
	vbaRec(&b, vbaIDDirTerm, nil)                           // dir terminator

	mods := parseVBADir(b.Bytes())
	if len(mods) != 1 {
		t.Fatalf("got %d modules, want 1: %+v", len(mods), mods)
	}
	if m := mods[0]; m.name != "Macro1" || m.streamName != "Macro1" || m.textOffset != 512 {
		t.Errorf("module = %+v, want {Macro1 Macro1 512}", m)
	}
}

// TestExtractVBADirect runs the whole extractVBA path against a real
// vbaProject.bin-style CFB: openStream -> readChain -> decompress -> parse
// dir -> decompress module source.
func TestExtractVBADirect(t *testing.T) {
	source := "Attribute VB_Name = \"Module1\"\r\nSub Hello()\r\n    MsgBox \"こんにちは マクロ\"\r\nEnd Sub\r\n"
	const cacheLen = 100
	module := append(make([]byte, cacheLen), compressVBALiteral([]byte(source))...)
	dir := compressVBALiteral(buildVBADir("Module1", cacheLen))

	f, err := parseCFB(buildCFB(map[string][]byte{"dir": dir, "Module1": module}))
	if err != nil {
		t.Fatalf("parseCFB: %v", err)
	}
	got := extractVBA(f)
	wantContains(t, got,
		"=== VBA Module: Module1 ===",
		`Attribute VB_Name = "Module1"`,
		"Sub Hello()",
		"こんにちは マクロ",
		"End Sub",
	)
	wantAbsent(t, got, "\r") // CRLF normalized to LF
}

// TestExtractXlsmMacro builds a genuine macro-enabled workbook (a real
// vbaProject.bin embedded in an .xlsx ZIP) and checks the VBA source comes
// out through Extract alongside the ordinary sheet text.
func TestExtractXlsmMacro(t *testing.T) {
	const source = "Attribute VB_Name = \"Module1\"\r\n" +
		"Sub Hello()\r\n" +
		"    MsgBox \"Hello from a macro\"\r\n" +
		"End Sub\r\n"
	const cacheLen = 10
	module := append(make([]byte, cacheLen), compressVBALiteral([]byte(source))...)
	dir := compressVBALiteral(buildVBADir("Module1", cacheLen))
	vbaBin := buildCFB(map[string][]byte{"dir": dir, "Module1": module})

	workbook := `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`
	rels := `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>`
	sheet := `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>cell text</t></is></c></row></sheetData></worksheet>`

	extractWant(t, zipBytes(map[string]string{
		"xl/workbook.xml":            workbook,
		"xl/_rels/workbook.xml.rels": rels,
		"xl/worksheets/sheet1.xml":   sheet,
		"xl/vbaProject.bin":          string(vbaBin),
	}),
		"cell text",
		"=== VBA Module: Module1 ===",
		"Sub Hello()",
		`MsgBox "Hello from a macro"`,
		"End Sub",
	)
}
