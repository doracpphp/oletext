// ooxml.go extracts text from the post-2007 Office Open XML formats
// (.docx, .xlsx, .pptx).
//
// Unlike the legacy .doc/.xls/.ppt formats -- which are OLE2 compound
// files parsed by cfb.go -- an OOXML document is a ZIP package ([ISO/IEC
// 29500], "Open Packaging Conventions") whose entries are XML parts. The
// package type is detected from the well-known part names inside the ZIP:
//
//	word/document.xml      -> WordprocessingML (.docx)
//	xl/workbook.xml        -> SpreadsheetML    (.xlsx)
//	ppt/presentation.xml   -> PresentationML   (.pptx)
//
// Large files: every XML part is streamed straight into encoding/xml from
// the ZIP entry's reader rather than being read whole into memory, and
// ExtractFile drives the package through zip.OpenReader, which reads only
// the central directory plus each part on demand. Peak memory therefore
// tracks the largest single part being decoded (and the accumulated output
// text), not the size of the whole document.
//
// Beyond ordinary body text, the extractor also gathers the "secondary"
// text Office hides in separate parts: footnotes, endnotes and comments
// (Word); cell comments, threaded comments and drawing/text-box shapes
// (Excel); and speaker notes and comments (PowerPoint). Elements are
// matched by their local name so the namespace prefix, which can vary, is
// irrelevant.

package oletext

import (
	"archive/zip"
	"bufio"
	"bytes"
	"encoding/xml"
	"errors"
	"fmt"
	"io"
	"path"
	"sort"
	"strconv"
	"strings"
)

// isZip reports whether data begins with a ZIP local-file, empty-archive
// or spanned-archive signature ("PK\x03\x04" / "PK\x05\x06" / "PK\x07\x08").
// Every OOXML package is a ZIP, so this distinguishes them from the OLE2
// compound files handled by cfb.go.
func isZip(data []byte) bool {
	if len(data) < 4 || data[0] != 'P' || data[1] != 'K' {
		return false
	}
	return (data[2] == 0x03 && data[3] == 0x04) ||
		(data[2] == 0x05 && data[3] == 0x06) ||
		(data[2] == 0x07 && data[3] == 0x08)
}

// extractOOXML parses an in-memory ZIP package and extracts its text.
func extractOOXML(data []byte) (string, error) {
	zr, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		return "", fmt.Errorf("not a valid ZIP/OOXML package: %w", err)
	}
	return extractOOXMLReader(zr)
}

// extractOOXMLReader detects the OOXML document type from the part names
// and dispatches to the matching extractor.
func extractOOXMLReader(zr *zip.Reader) (string, error) {
	files := make(map[string]*zip.File, len(zr.File))
	for _, zf := range zr.File {
		files[zf.Name] = zf
	}
	switch {
	case files["word/document.xml"] != nil:
		return extractDocx(files)
	case files["xl/workbook.xml"] != nil:
		return extractXlsx(files)
	case files["ppt/presentation.xml"] != nil:
		return extractPptx(files)
	default:
		return "", errors.New("unrecognized OOXML package (no Word/Excel/PowerPoint part found)")
	}
}

// ---- WordprocessingML (.docx) ----

// extractDocx extracts the main document body followed by the auxiliary
// parts (headers, footers, footnotes, endnotes and comments). Shape and
// text-box text lives inline in word/document.xml (and the headers), so it
// is captured along with ordinary runs.
func extractDocx(files map[string]*zip.File) (string, error) {
	var parts []string
	if t := extractPartText(files["word/document.xml"]); strings.TrimSpace(t) != "" {
		parts = append(parts, t)
	}
	for _, name := range sortedParts(files, isDocxAuxPart) {
		if t := extractPartText(files[name]); strings.TrimSpace(t) != "" {
			parts = append(parts, t)
		}
	}
	parts = append(parts, ooxmlExtras(files, "word/vbaProject.bin")...)
	return joinParts(parts), nil
}

// isDocxAuxPart reports whether a part is a WordprocessingML fragment
// carried alongside word/document.xml (footnotes, endnotes, comments,
// headers and footers).
func isDocxAuxPart(name string) bool {
	switch name {
	case "word/footnotes.xml", "word/endnotes.xml", "word/comments.xml":
		return true
	}
	if !strings.HasSuffix(name, ".xml") {
		return false
	}
	return strings.HasPrefix(name, "word/header") || strings.HasPrefix(name, "word/footer")
}

// ---- PresentationML (.pptx) ----

// extractPptx extracts each slide's text (including shapes and text boxes,
// which are ordinary DrawingML <a:t> runs) in presentation order, then the
// speaker notes and comments.
func extractPptx(files map[string]*zip.File) (string, error) {
	var parts []string
	for _, name := range pptxSlideOrder(files) {
		if t := extractPartText(files[name]); strings.TrimSpace(t) != "" {
			parts = append(parts, t)
		}
	}
	for _, name := range sortedParts(files, isPptxNotePart) {
		if t := extractPartText(files[name]); strings.TrimSpace(t) != "" {
			parts = append(parts, "=== Notes ===\n"+t)
		}
	}
	// Comment text sits in <p:text> (classic) or <a:t> (modern threaded).
	for _, name := range sortedParts(files, isPptxCommentPart) {
		if t := extractPartText(files[name], "text"); strings.TrimSpace(t) != "" {
			parts = append(parts, "=== Comments ===\n"+t)
		}
	}
	// Slide masters and layouts carry placeholder/prompt text and any text
	// the author added to the template (logos captions, fixed footers).
	for _, name := range sortedParts(files, isPptxMasterPart) {
		if t := extractPartText(files[name]); strings.TrimSpace(t) != "" {
			parts = append(parts, "=== Layout ===\n"+t)
		}
	}
	parts = append(parts, ooxmlExtras(files, "ppt/vbaProject.bin")...)
	return joinParts(parts), nil
}

func isPptxNotePart(name string) bool {
	return strings.HasPrefix(name, "ppt/notesSlides/notesSlide") && strings.HasSuffix(name, ".xml")
}

func isPptxCommentPart(name string) bool {
	return strings.HasPrefix(name, "ppt/comments/") && strings.HasSuffix(name, ".xml")
}

func isPptxMasterPart(name string) bool {
	if !strings.HasSuffix(name, ".xml") {
		return false
	}
	return strings.HasPrefix(name, "ppt/slideMasters/slideMaster") ||
		strings.HasPrefix(name, "ppt/slideLayouts/slideLayout")
}

// pptxSlideOrder returns the slide part names in presentation order,
// resolved through ppt/presentation.xml and its relationships. It falls
// back to the numeric order of ppt/slides/slideN.xml when those parts are
// unavailable.
func pptxSlideOrder(files map[string]*zip.File) []string {
	relMap := parseRels(files["ppt/_rels/presentation.xml.rels"])
	var order []string
	forEachStartElement(files["ppt/presentation.xml"], func(se xml.StartElement) {
		if se.Name.Local == "sldId" {
			if target, ok := relMap[relID(se.Attr)]; ok {
				order = append(order, resolveTarget("ppt", target))
			}
		}
	})
	if len(order) > 0 {
		return order
	}
	return sortedParts(files, func(name string) bool {
		return strings.HasPrefix(name, "ppt/slides/slide") && strings.HasSuffix(name, ".xml")
	})
}

// ---- SpreadsheetML (.xlsx) ----

// sheetRef pairs a worksheet's display name with its part name.
type sheetRef struct {
	name string
	part string
}

// extractXlsx extracts every worksheet's cells (in workbook order), then
// the cell comments and the text of drawing shapes / text boxes. Cells in
// a row are tab-separated and each sheet is introduced with a header.
func extractXlsx(files map[string]*zip.File) (string, error) {
	sst := xlsxSharedStrings(files["xl/sharedStrings.xml"])

	var parts []string
	for _, s := range xlsxSheetOrder(files) {
		var b strings.Builder
		fmt.Fprintf(&b, "=== Sheet: %s ===\n", s.name)
		b.WriteString(xlsxSheetText(files[s.part], sst))
		parts = append(parts, strings.TrimRight(b.String(), "\n"))
	}
	// Classic comment text is <text><r><t>; threaded comments store it
	// directly in <text>. Capturing both <t> and <text> covers each.
	for _, name := range sortedParts(files, isXlsxCommentPart) {
		if t := extractPartText(files[name], "text"); strings.TrimSpace(t) != "" {
			parts = append(parts, "=== Comments ===\n"+t)
		}
	}
	for _, name := range sortedParts(files, isXlsxDrawingPart) {
		if t := extractPartText(files[name]); strings.TrimSpace(t) != "" {
			parts = append(parts, "=== Shapes ===\n"+t)
		}
	}
	parts = append(parts, ooxmlExtras(files, "xl/vbaProject.bin")...)
	return joinSections(parts), nil
}

func isXlsxCommentPart(name string) bool {
	if !strings.HasSuffix(name, ".xml") {
		return false
	}
	return strings.HasPrefix(name, "xl/comments") || strings.HasPrefix(name, "xl/threadedComments/")
}

func isXlsxDrawingPart(name string) bool {
	return strings.HasPrefix(name, "xl/drawings/drawing") && strings.HasSuffix(name, ".xml")
}

// ---- text shared by all OOXML formats: charts, properties, macros ----

// ooxmlExtras gathers the text every OOXML package can carry regardless of
// application: chart and SmartArt diagram labels (DrawingML <a:t>), the
// core document properties, and -- for a macro-enabled document -- the VBA
// macro source in vbaPart (xl/ word/ or ppt/ vbaProject.bin).
func ooxmlExtras(files map[string]*zip.File, vbaPart string) []string {
	var extra []string
	for _, name := range sortedParts(files, isChartOrDiagramPart) {
		if t := extractPartText(files[name]); strings.TrimSpace(t) != "" {
			extra = append(extra, "=== Chart ===\n"+t)
		}
	}
	if t := corePropertiesText(files["docProps/core.xml"]); strings.TrimSpace(t) != "" {
		extra = append(extra, "=== Properties ===\n"+t)
	}
	if m := ooxmlMacroText(files[vbaPart]); strings.TrimSpace(m) != "" {
		extra = append(extra, "=== Macros ===\n"+strings.TrimRight(m, "\n"))
	}
	return extra
}

// isChartOrDiagramPart matches the DrawingML chart and SmartArt diagram
// data parts of any of the three formats (word/xl/ppt charts and diagrams).
func isChartOrDiagramPart(name string) bool {
	if !strings.HasSuffix(name, ".xml") {
		return false
	}
	return strings.Contains(name, "/charts/chart") || strings.Contains(name, "/diagrams/data")
}

// ooxmlMacroText extracts the VBA macro source from an embedded
// vbaProject.bin part. The part is itself an OLE2 compound file, so it is
// parsed by parseCFB and handed to the shared extractVBA used for the
// legacy binary formats. The whole part is read into memory (macro
// projects are small and need random access).
func ooxmlMacroText(zf *zip.File) string {
	bin := readPartBytes(zf)
	if len(bin) == 0 {
		return ""
	}
	f, err := parseCFB(bin)
	if err != nil {
		return ""
	}
	return extractVBA(f)
}

// corePropTags are the docProps/core.xml elements that carry human-authored
// text. Timestamps, revision and version numbers are deliberately excluded.
var corePropTags = map[string]bool{
	"title": true, "subject": true, "creator": true, "keywords": true,
	"description": true, "lastModifiedBy": true, "category": true, "contentStatus": true,
}

// corePropertiesText returns the meaningful text of docProps/core.xml: the
// title, subject, author, keywords, description and similar fields, joined
// by newlines.
func corePropertiesText(zf *zip.File) string {
	dec, closer := newPartDecoder(zf)
	if dec == nil {
		return ""
	}
	defer closer.Close()
	var parts []string
	capture := false
	for {
		tok, err := dec.Token()
		if err != nil {
			break
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if corePropTags[t.Name.Local] {
				capture = true
			}
		case xml.EndElement:
			if corePropTags[t.Name.Local] {
				capture = false
			}
		case xml.CharData:
			if capture {
				if s := strings.TrimSpace(string(t)); s != "" {
					parts = append(parts, s)
				}
			}
		}
	}
	return strings.Join(parts, "\n")
}

// readPartBytes returns the full decompressed contents of a ZIP part, or
// nil if it is missing or unreadable. Used only for the small binary
// vbaProject.bin part, which needs random access.
func readPartBytes(zf *zip.File) []byte {
	rc := openPart(zf)
	if rc == nil {
		return nil
	}
	defer rc.Close()
	data, err := io.ReadAll(rc)
	if err != nil {
		return nil
	}
	return data
}

// xlsxSheetOrder returns the worksheets in workbook order with their
// names, resolved through xl/workbook.xml and its relationships. It falls
// back to the numeric order of xl/worksheets/sheetN.xml when those parts
// are unavailable.
func xlsxSheetOrder(files map[string]*zip.File) []sheetRef {
	relMap := parseRels(files["xl/_rels/workbook.xml.rels"])
	var refs []sheetRef
	forEachStartElement(files["xl/workbook.xml"], func(se xml.StartElement) {
		if se.Name.Local != "sheet" {
			return
		}
		name := ""
		for _, a := range se.Attr {
			if a.Name.Local == "name" {
				name = a.Value
			}
		}
		if target, ok := relMap[relID(se.Attr)]; ok {
			refs = append(refs, sheetRef{name, resolveTarget("xl", target)})
		}
	})
	if len(refs) > 0 {
		return refs
	}

	names := sortedParts(files, func(name string) bool {
		return strings.HasPrefix(name, "xl/worksheets/sheet") && strings.HasSuffix(name, ".xml")
	})
	refs = make([]sheetRef, len(names))
	for i, n := range names {
		refs[i] = sheetRef{fmt.Sprintf("Sheet%d", i+1), n}
	}
	return refs
}

// xlsxSharedStrings reads xl/sharedStrings.xml into a slice indexed by the
// shared-string index a cell carries. Each <si> string item is the
// concatenation of its <t> runs; phonetic guides (<rPh>) are skipped.
func xlsxSharedStrings(zf *zip.File) []string {
	dec, closer := newPartDecoder(zf)
	if dec == nil {
		return nil
	}
	defer closer.Close()

	var sst []string
	var cur strings.Builder
	inSI, inT, skip := false, false, 0
	for {
		tok, err := dec.Token()
		if err != nil {
			break
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "si":
				inSI, inT, skip = true, false, 0
				cur.Reset()
			case "rPh":
				skip++
			case "t":
				inT = true
			}
		case xml.EndElement:
			switch t.Name.Local {
			case "t":
				inT = false
			case "rPh":
				if skip > 0 {
					skip--
				}
			case "si":
				inSI = false
				sst = append(sst, cur.String())
			}
		case xml.CharData:
			if inSI && inT && skip == 0 {
				cur.Write(t)
			}
		}
	}
	return sst
}

// xlsxSheetText extracts the non-empty cells of a worksheet. A cell's text
// comes from its <v> element (a number, formula-string result, or a shared
// string index when t="s") or from <is><t> when t="inlineStr". The print
// header/footer text in <headerFooter> (whose &L/&C/&R section codes are
// stripped) is appended after the cell grid.
func xlsxSheetText(zf *zip.File, sst []string) string {
	dec, closer := newPartDecoder(zf)
	if dec == nil {
		return ""
	}
	defer closer.Close()

	var sb strings.Builder
	var row []string
	var cell strings.Builder
	cellType := ""
	capture := false
	var hf strings.Builder
	hfCapture := false
	var hfLines []string
	seenHF := map[string]bool{}
	for {
		tok, err := dec.Token()
		if err != nil {
			break
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "row":
				row = row[:0]
			case "c":
				cellType = ""
				for _, a := range t.Attr {
					if a.Name.Local == "t" {
						cellType = a.Value
					}
				}
				cell.Reset()
			case "v", "t":
				capture = true
			case "oddHeader", "oddFooter", "evenHeader", "evenFooter", "firstHeader", "firstFooter":
				hfCapture = true
				hf.Reset()
			}
		case xml.EndElement:
			switch t.Name.Local {
			case "v", "t":
				capture = false
			case "c":
				text := cell.String()
				if cellType == "s" {
					if idx, err := strconv.Atoi(strings.TrimSpace(text)); err == nil && idx >= 0 && idx < len(sst) {
						text = sst[idx]
					} else {
						text = ""
					}
				}
				if text != "" {
					row = append(row, text)
				}
			case "row":
				if len(row) > 0 {
					sb.WriteString(strings.Join(row, "\t"))
					sb.WriteByte('\n')
				}
			case "oddHeader", "oddFooter", "evenHeader", "evenFooter", "firstHeader", "firstFooter":
				hfCapture = false
				if s := stripHeaderFooterCodes(hf.String()); s != "" && !seenHF[s] {
					seenHF[s] = true
					hfLines = append(hfLines, s)
				}
			}
		case xml.CharData:
			if capture {
				cell.Write(t)
			} else if hfCapture {
				hf.Write(t)
			}
		}
	}
	for _, line := range hfLines {
		sb.WriteString(line)
		sb.WriteByte('\n')
	}
	return sb.String()
}

// ---- shared OOXML helpers ----

// extractXMLText streams a WordprocessingML/DrawingML part and pulls out
// its running text. By default the text-bearing elements are <w:t>/<a:t>;
// extra local names (such as "text" for comment parts) can be added. A
// <w:p>/<a:p> ends a paragraph, <w:tab> becomes a tab and <w:br>/<w:cr> a
// newline. The <mc:Fallback> branch of an mc:AlternateContent block is
// skipped so its content is not emitted twice (once per alternative).
func extractXMLText(r io.Reader, textTags ...string) string {
	capTag := map[string]bool{"t": true}
	for _, tg := range textTags {
		capTag[tg] = true
	}
	dec := xml.NewDecoder(bomReader(r))
	dec.Strict = false

	var sb strings.Builder
	depth := 0   // current element nesting depth
	capAt := -1  // depth at which text capture began (-1 = not capturing)
	skipAt := -1 // depth at which a skipped subtree began (-1 = none)
	for {
		tok, err := dec.Token()
		if err != nil {
			break
		}
		switch t := tok.(type) {
		case xml.StartElement:
			depth++
			if skipAt >= 0 {
				continue
			}
			switch {
			case t.Name.Local == "Fallback":
				skipAt = depth
			case capTag[t.Name.Local]:
				if capAt < 0 {
					capAt = depth
				}
			case t.Name.Local == "tab":
				sb.WriteByte('\t')
			case t.Name.Local == "br", t.Name.Local == "cr":
				sb.WriteByte('\n')
			}
		case xml.EndElement:
			if skipAt >= 0 {
				if depth == skipAt {
					skipAt = -1
				}
				depth--
				continue
			}
			if capAt >= 0 && depth == capAt {
				capAt = -1
			}
			if t.Name.Local == "p" && capAt < 0 {
				sb.WriteByte('\n')
			}
			depth--
		case xml.CharData:
			if skipAt < 0 && capAt >= 0 {
				sb.Write(t)
			}
		}
	}
	return strings.Trim(sb.String(), "\n")
}

// extractPartText opens a ZIP part and extracts its text, streaming it
// straight into the decoder. Returns "" for a missing or unreadable part.
func extractPartText(zf *zip.File, textTags ...string) string {
	rc := openPart(zf)
	if rc == nil {
		return ""
	}
	defer rc.Close()
	return extractXMLText(rc, textTags...)
}

// parseRels parses an Open Packaging Conventions relationships part into a
// map from relationship Id to Target.
func parseRels(zf *zip.File) map[string]string {
	m := map[string]string{}
	forEachStartElement(zf, func(se xml.StartElement) {
		if se.Name.Local != "Relationship" {
			return
		}
		var id, target string
		for _, a := range se.Attr {
			switch a.Name.Local {
			case "Id":
				id = a.Value
			case "Target":
				target = a.Value
			}
		}
		if id != "" && target != "" {
			m[id] = target
		}
	})
	return m
}

// relID returns the value of the r:id attribute (the relationship-namespace
// "id"), distinguishing it from a plain unqualified "id" attribute that may
// sit on the same element (e.g. <p:sldId id="256" r:id="rId2">).
func relID(attrs []xml.Attr) string {
	for _, a := range attrs {
		if a.Name.Local == "id" && a.Name.Space != "" {
			return a.Value
		}
	}
	return ""
}

// resolveTarget resolves a relationship Target against the directory of the
// part that owns the relationships. An absolute target ("/foo") is taken
// relative to the package root.
func resolveTarget(base, target string) string {
	if rooted, ok := strings.CutPrefix(target, "/"); ok {
		return rooted
	}
	return path.Join(base, target)
}

// forEachStartElement streams a ZIP part and invokes fn for every start
// element. It does nothing for a missing part and stops at the first error
// or EOF.
func forEachStartElement(zf *zip.File, fn func(xml.StartElement)) {
	dec, closer := newPartDecoder(zf)
	if dec == nil {
		return
	}
	defer closer.Close()
	for {
		tok, err := dec.Token()
		if err != nil {
			return
		}
		if se, ok := tok.(xml.StartElement); ok {
			fn(se)
		}
	}
}

// newPartDecoder opens a ZIP part for streaming XML decoding. It returns a
// nil decoder for a missing or unreadable part; otherwise the caller must
// Close the returned io.Closer.
func newPartDecoder(zf *zip.File) (*xml.Decoder, io.Closer) {
	rc := openPart(zf)
	if rc == nil {
		return nil, nil
	}
	dec := xml.NewDecoder(bomReader(rc))
	dec.Strict = false
	return dec, rc
}

// openPart opens a ZIP entry for reading, or returns nil if it is missing
// or cannot be opened.
func openPart(zf *zip.File) io.ReadCloser {
	if zf == nil {
		return nil
	}
	rc, err := zf.Open()
	if err != nil {
		return nil
	}
	return rc
}

// bomReader wraps r so that a leading UTF-8 byte-order mark -- which some
// Office parts carry and which encoding/xml will not skip on its own -- is
// discarded. The bufio wrapper also gives the decoder buffered reads.
func bomReader(r io.Reader) io.Reader {
	br := bufio.NewReader(r)
	if b, err := br.Peek(3); err == nil && b[0] == 0xEF && b[1] == 0xBB && b[2] == 0xBF {
		br.Discard(3)
	}
	return br
}

// sortedParts returns the names of the parts matching pred, ordered by the
// trailing integer in their base name.
func sortedParts(files map[string]*zip.File, pred func(string) bool) []string {
	var names []string
	for name := range files {
		if pred(name) {
			names = append(names, name)
		}
	}
	sortNumeric(names)
	return names
}

// joinParts joins extracted parts with single blank-line-free newlines,
// terminating with a newline. It returns "" when there is nothing.
func joinParts(parts []string) string {
	if len(parts) == 0 {
		return ""
	}
	return strings.Join(parts, "\n") + "\n"
}

// joinSections joins labelled sections (sheets, comments, shapes) with a
// blank line between them, terminating with a newline.
func joinSections(parts []string) string {
	if len(parts) == 0 {
		return ""
	}
	return strings.Join(parts, "\n\n") + "\n"
}

// sortNumeric sorts part names by the trailing integer in their base name
// (sheet2.xml before sheet10.xml).
func sortNumeric(names []string) {
	sort.Slice(names, func(i, j int) bool {
		ni, nj := partNum(names[i]), partNum(names[j])
		if ni != nj {
			return ni < nj
		}
		return names[i] < names[j]
	})
}

// partNum extracts the trailing integer of a part's base name, e.g.
// "ppt/slides/slide12.xml" -> 12. It returns 0 when there is no number.
func partNum(name string) int {
	b := path.Base(name)
	b = strings.TrimSuffix(b, path.Ext(b))
	i := len(b)
	for i > 0 && b[i-1] >= '0' && b[i-1] <= '9' {
		i--
	}
	n, _ := strconv.Atoi(b[i:])
	return n
}
