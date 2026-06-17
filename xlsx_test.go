package oletext

import (
	"strings"
	"testing"
)

// xlsx tests. Body, cell comment, print header/footer and document
// properties run against the real testdata/sample.xlsx; the phonetic-guide,
// threaded-comment, drawing-shape and header/footer-code cases use in-memory
// packages because openpyxl cannot easily author them.

func TestXlsxBody(t *testing.T) {
	got := extractFileWant(t, "testdata/sample.xlsx",
		"=== Sheet: Numbers ===",
		"Label\tValue",
		"Answer\t42",
		"日本語セル",
		"=== Sheet: Second ===",
		"On the second sheet",
	)
	if strings.Index(got, "Numbers") > strings.Index(got, "Second") {
		t.Errorf("sheets out of order:\n%s", got)
	}
}

func TestXlsxComment(t *testing.T) {
	extractFileWant(t, "testdata/sample.xlsx", "A cell comment.")
}

func TestXlsxHeaderFooter(t *testing.T) {
	extractFileWant(t, "testdata/sample.xlsx",
		"Printed Report Header",
		"Sheet Footer Note",
	)
}

func TestXlsxProperties(t *testing.T) {
	extractFileWant(t, "testdata/sample.xlsx",
		"Sheet Property Title",
		"Sheet Author",
	)
}

// TestXlsxPhonetic checks the phonetic guide (<rPh>) is skipped while the
// base shared-string text is kept.
func TestXlsxPhonetic(t *testing.T) {
	workbook := `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`
	rels := `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>`
	shared := `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t>Name</t><rPh sb="0" eb="1"><t>ナマエ</t></rPh></si></sst>`
	sheet := `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData></worksheet>`
	got := extractWant(t, zipBytes(map[string]string{
		"xl/workbook.xml":            workbook,
		"xl/_rels/workbook.xml.rels": rels,
		"xl/sharedStrings.xml":       shared,
		"xl/worksheets/sheet1.xml":   sheet,
	}), "Name")
	wantAbsent(t, got, "ナマエ")
}

// TestXlsxThreadedComment checks the modern threaded-comment kind (text held
// directly in <text>).
func TestXlsxThreadedComment(t *testing.T) {
	workbook := `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`
	rels := `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>`
	sheet := `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>cell</t></is></c></row></sheetData></worksheet>`
	threaded := `<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"><threadedComment ref="C3" id="{x}"><text>A threaded comment.</text></threadedComment></ThreadedComments>`
	extractWant(t, zipBytes(map[string]string{
		"xl/workbook.xml":            workbook,
		"xl/_rels/workbook.xml.rels": rels,
		"xl/worksheets/sheet1.xml":   sheet,
		"xl/threadedComments/threadedComment1.xml": threaded,
	}), "cell", "A threaded comment.")
}

// TestXlsxDrawingShape checks text-box shape text on the grid (DrawingML).
func TestXlsxDrawingShape(t *testing.T) {
	workbook := `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`
	rels := `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>`
	sheet := `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>grid cell</t></is></c></row></sheetData></worksheet>`
	drawing := `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><xdr:sp><xdr:txBody><a:p><a:r><a:t>Shape on the grid.</a:t></a:r></a:p></xdr:txBody></xdr:sp></xdr:wsDr>`
	extractWant(t, zipBytes(map[string]string{
		"xl/workbook.xml":            workbook,
		"xl/_rels/workbook.xml.rels": rels,
		"xl/worksheets/sheet1.xml":   sheet,
		"xl/drawings/drawing1.xml":   drawing,
	}), "grid cell", "Shape on the grid.")
}

// TestXlsxHeaderFooterCodes checks the &L/&C/&R section codes and font codes
// are stripped from print header/footer text.
func TestXlsxHeaderFooterCodes(t *testing.T) {
	workbook := `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`
	rels := `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>`
	sheet := `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
		`<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>grid cell</t></is></c></row></sheetData>` +
		`<headerFooter><oddHeader>&L&"Arial,Bold"&14LeftHead&CCenterHead&RRightHead</oddHeader>` +
		`<oddFooter>&CConfidential &P</oddFooter></headerFooter></worksheet>`
	got := extractWant(t, zipBytes(map[string]string{
		"xl/workbook.xml":            workbook,
		"xl/_rels/workbook.xml.rels": rels,
		"xl/worksheets/sheet1.xml":   sheet,
	}), "grid cell", "LeftHead", "CenterHead", "RightHead", "Confidential")
	wantAbsent(t, got, "&L", "&C", "&14", "Arial")
}
