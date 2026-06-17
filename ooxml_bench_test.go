package oletext

import (
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

// These benchmarks build large in-memory .docx/.xlsx/.pptx packages and
// measure extraction throughput, so the streaming design (which never
// holds a whole part in memory beyond the decoder window) can be checked
// against documents far larger than any committed sample. b.SetBytes makes
// the reported MB/s the compressed package size processed per second. The
// zipBytes packer lives in testutil_test.go.

// bigDocx builds a .docx whose body holds nPara paragraphs.
func bigDocx(nPara int) []byte {
	var b strings.Builder
	b.WriteString(`<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>`)
	for i := 0; i < nPara; i++ {
		fmt.Fprintf(&b, `<w:p><w:r><w:t>Paragraph %d: the quick brown fox jumps over the lazy dog.</w:t></w:r></w:p>`, i)
	}
	b.WriteString(`</w:body></w:document>`)
	return zipBytes(map[string]string{"word/document.xml": b.String()})
}

// bigXlsx builds a .xlsx with one sheet of nRow rows by nCol columns,
// alternating shared-string and numeric cells.
func bigXlsx(nRow, nCol int) []byte {
	var sst strings.Builder
	sst.WriteString(`<?xml version="1.0"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">`)
	for i := 0; i < 64; i++ {
		fmt.Fprintf(&sst, `<si><t>shared string number %d</t></si>`, i)
	}
	sst.WriteString(`</sst>`)

	var sheet strings.Builder
	sheet.WriteString(`<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>`)
	for r := 1; r <= nRow; r++ {
		fmt.Fprintf(&sheet, `<row r="%d">`, r)
		for c := 0; c < nCol; c++ {
			if c%2 == 0 {
				fmt.Fprintf(&sheet, `<c t="s"><v>%d</v></c>`, (r+c)%64)
			} else {
				fmt.Fprintf(&sheet, `<c><v>%d.%d</v></c>`, r, c)
			}
		}
		sheet.WriteString(`</row>`)
	}
	sheet.WriteString(`</sheetData></worksheet>`)

	workbook := `<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Big" sheetId="1" r:id="rId1"/></sheets></workbook>`
	rels := `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>`
	return zipBytes(map[string]string{
		"xl/workbook.xml":            workbook,
		"xl/_rels/workbook.xml.rels": rels,
		"xl/sharedStrings.xml":       sst.String(),
		"xl/worksheets/sheet1.xml":   sheet.String(),
	})
}

// bigPptx builds a .pptx with nSlide slides, each carrying a text box.
func bigPptx(nSlide int) []byte {
	parts := map[string]string{}
	var pres, rels strings.Builder
	pres.WriteString(`<?xml version="1.0"?><p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst>`)
	rels.WriteString(`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`)
	for i := 1; i <= nSlide; i++ {
		fmt.Fprintf(&pres, `<p:sldId id="%d" r:id="rId%d"/>`, 255+i, i)
		fmt.Fprintf(&rels, `<Relationship Id="rId%d" Target="slides/slide%d.xml"/>`, i, i)
		parts[fmt.Sprintf("ppt/slides/slide%d.xml", i)] = fmt.Sprintf(
			`<?xml version="1.0"?><p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>Slide %d body text content for benchmarking.</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`, i)
	}
	pres.WriteString(`</p:sldIdLst></p:presentation>`)
	rels.WriteString(`</Relationships>`)
	parts["ppt/presentation.xml"] = pres.String()
	parts["ppt/_rels/presentation.xml.rels"] = rels.String()
	return zipBytes(parts)
}

func benchExtract(b *testing.B, data []byte) {
	b.SetBytes(int64(len(data)))
	b.ReportAllocs()
	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		if _, err := Extract(data); err != nil {
			b.Fatal(err)
		}
	}
}

func BenchmarkExtractDocxLarge(b *testing.B) { benchExtract(b, bigDocx(50000)) }
func BenchmarkExtractXlsxLarge(b *testing.B) { benchExtract(b, bigXlsx(50000, 10)) }
func BenchmarkExtractPptxLarge(b *testing.B) { benchExtract(b, bigPptx(2000)) }

// BenchmarkExtractFileXlsxLarge measures the on-disk streaming path
// (zip.OpenReader) that ExtractFile uses for real files.
func BenchmarkExtractFileXlsxLarge(b *testing.B) {
	data := bigXlsx(50000, 10)
	path := filepath.Join(b.TempDir(), "big.xlsx")
	if err := os.WriteFile(path, data, 0o644); err != nil {
		b.Fatal(err)
	}
	b.SetBytes(int64(len(data)))
	b.ReportAllocs()
	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		if _, err := ExtractFile(path); err != nil {
			b.Fatal(err)
		}
	}
}
