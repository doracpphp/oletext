package oletext

import (
	"bytes"
	"testing"
)

// Fuzz targets for robustness: on arbitrary or mutated input the parsers
// must never panic (out-of-range slice, huge allocation, unbounded
// recursion). They check the code returns rather than crashing. The seed
// helpers (addBinarySeeds, streamSeeds, splitBytes) live in testutil_test.go.

// FuzzExtract drives the whole pipeline on mutated real files and arbitrary
// bytes.
func FuzzExtract(f *testing.F) {
	addBinarySeeds(f)
	f.Add([]byte{})
	f.Add([]byte("not an office file"))
	f.Fuzz(func(t *testing.T, data []byte) {
		_, _ = Extract(data)
	})
}

// FuzzExtractOOXMLPart wraps fuzzed bytes as the body part of an otherwise
// valid .docx/.xlsx/.pptx, reaching the OOXML token/cell/shared-string
// handlers with malformed XML.
func FuzzExtractOOXMLPart(f *testing.F) {
	f.Add([]byte(`<w:document xmlns:w="x"><w:body><w:p><w:t>hi</w:t></w:p></w:body></w:document>`))
	f.Add([]byte("<not-closed"))
	f.Add([]byte("\xef\xbb\xbf<a><b/></a>"))
	f.Add([]byte{})

	const wbRels = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>`
	const presRels = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Target="slides/slide1.xml"/></Relationships>`
	const workbook = `<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`
	const pres = `<p:presentation xmlns:p="x" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst></p:presentation>`

	f.Fuzz(func(t *testing.T, data []byte) {
		s := string(data)
		_, _ = Extract(zipBytes(map[string]string{"word/document.xml": s}))
		_, _ = Extract(zipBytes(map[string]string{
			"xl/workbook.xml":            workbook,
			"xl/_rels/workbook.xml.rels": wbRels,
			"xl/sharedStrings.xml":       s,
			"xl/worksheets/sheet1.xml":   s,
		}))
		_, _ = Extract(zipBytes(map[string]string{
			"ppt/presentation.xml":            pres,
			"ppt/_rels/presentation.xml.rels": presRels,
			"ppt/slides/slide1.xml":           s,
			"ppt/notesSlides/notesSlide1.xml": s,
			"ppt/comments/comment1.xml":       s,
		}))
	})
}

// FuzzOfficeXMLText fuzzes the shared WordprocessingML/DrawingML text walker.
func FuzzOfficeXMLText(f *testing.F) {
	f.Add([]byte(`<a><t>x</t><p/><tab/><br/></a>`))
	f.Add([]byte(`<mc:AlternateContent><mc:Fallback><t>dup</t></mc:Fallback></mc:AlternateContent>`))
	f.Add([]byte{})
	f.Fuzz(func(t *testing.T, data []byte) {
		_ = extractXMLText(bytes.NewReader(data), "text")
	})
}

// FuzzWalkPptRecords fuzzes the recursive PowerPoint record walk.
func FuzzWalkPptRecords(f *testing.F) {
	for _, s := range streamSeeds("PowerPoint Document") {
		f.Add(s)
	}
	f.Add([]byte{})
	f.Fuzz(func(t *testing.T, data []byte) {
		var parts []string
		encrypted := false
		walkPptRecords(data, 0, &parts, &encrypted)
	})
}

// FuzzDecodePieceTable fuzzes the Word piece-table decoder.
func FuzzDecodePieceTable(f *testing.F) {
	f.Add([]byte{}, []byte{})
	f.Add([]byte{0, 0, 0, 0, 10, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}, []byte("hello world"))
	f.Fuzz(func(t *testing.T, plc, wd []byte) {
		_, _ = decodePieceTable(plc, wd)
	})
}

// FuzzXlsHandlers feeds arbitrary bytes to every BIFF record handler in both
// the BIFF8 and BIFF5 paths.
func FuzzXlsHandlers(f *testing.F) {
	f.Add([]byte{})
	f.Add([]byte{0, 0, 0, 12, 11, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 'M', 'y'})
	f.Fuzz(func(t *testing.T, d []byte) {
		for _, biff5 := range []bool{false, true} {
			x := &xlsExtractor{curRow: -1, sheetIdx: -1, atLineStart: true, biff5: biff5}
			x.hasPendingString = true
			x.onBOF(d)
			x.onEOF()
			x.onBoundSheet(d)
			x.onSST([][]byte{d})
			x.onLabelSst(d)
			x.onLabel(d)
			x.onNumber(d)
			x.onRK(d)
			x.onMulRk(d)
			x.onFormula(d)
			x.onString(d)
			x.onTxO(d, [][]byte{d})
			x.onHeaderFooter(d)
			x.onLbl(d)
			x.onSeriesText(d)
			x.onHLink(d)
		}
	})
}

// FuzzSplitBiffRecords fuzzes the BIFF record slicer.
func FuzzSplitBiffRecords(f *testing.F) {
	for _, s := range streamSeeds("Workbook") {
		f.Add(s)
	}
	f.Add([]byte{})
	f.Fuzz(func(t *testing.T, data []byte) {
		_ = splitBiffRecords(data)
	})
}

// FuzzSegReader fuzzes the cross-Continue string readers.
func FuzzSegReader(f *testing.F) {
	f.Add([]byte{}, uint8(0), uint8(0), uint16(0), false)
	f.Add([]byte{5, 0, 1, 'a', 0, 'b', 0}, uint8(3), uint8(1), uint16(2), true)
	f.Fuzz(func(t *testing.T, d []byte, splitAt, cchSize uint8, cch uint16, high bool) {
		segs := splitBytes(d, int(splitAt))
		readXLString(newSegReader(segs), 1+int(cchSize%2))
		newSegReader(segs).readChars(int(cch), high)
		r := newSegReader(segs)
		r.skip(int(cch))
		r.readU32()
	})
}

// FuzzHyperlinkAndCodes fuzzes the hyperlink string reader and the
// header/footer code stripper.
func FuzzHyperlinkAndCodes(f *testing.F) {
	f.Add([]byte{}, uint16(0))
	f.Add([]byte("&C&\"Font\"&12Title&P&Kff0000"), uint16(0))
	f.Fuzz(func(t *testing.T, d []byte, pos uint16) {
		readHyperlinkString(d, int(pos))
		_ = isURLMoniker(d)
		_ = decodeUTF16(d)
		_ = stripHeaderFooterCodes(string(d))
		_ = cleanAuxText(string(d))
		_ = cleanDocText(string(d))
		_ = cleanPptText(string(d))
	})
}

// FuzzDecompressVBA checks the MS-OVBA decompressor never panics and stays
// within its output cap.
func FuzzDecompressVBA(f *testing.F) {
	f.Add([]byte{0x01, 0x05, 0xB0, 0x08, 'A', 'B', 'C', 0x00, 0x20})
	f.Add([]byte{0x01, 0x02, 0x30, 'X', 'Y', 'Z'})
	f.Add([]byte{})
	f.Fuzz(func(t *testing.T, data []byte) {
		if out := decompressVBA(data); len(out) > maxVBAOutput {
			t.Fatalf("output %d exceeds cap %d", len(out), maxVBAOutput)
		}
	})
}
