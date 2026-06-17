// xls.go extracts text from Excel binary files (.xls).
//
// Referenced sections of [MS-XLS]:
//   - 2.1.4 Record (type:2, size:2, payload)
//   - 2.4.265 SST / 2.4.58 Continue (when a string spans a Continue record,
//     the continued part starts with a fresh fHighByte flag byte)
//   - 2.5.293 XLUnicodeRichExtendedString, 2.5.294 XLUnicodeString
//   - 2.4.149 LabelSst, 2.4.148 Label, 2.4.180 Number, 2.4.220 RK,
//     2.4.175 MulRk, 2.4.127 Formula, 2.4.268 String, 2.4.28 BoundSheet8
//   - 2.4.324 TxO (text of shapes, text boxes and cell comments; the text
//     itself is stored in the Continue records that follow the TxO)
//
// A workbook also holds text outside the cell grid and the drawing layer.
// These records are extracted too:
//   - 2.4.137 Header / 2.4.136 Footer (page header/footer text, carrying
//     page-setup format codes that stripHeaderFooterCodes removes)
//   - 2.4.150 Lbl (defined names / named ranges)
//   - 2.4.252 SeriesText (chart titles, axis titles and series names, found
//     in the chart substreams)
//   - 2.4.323 HLink (hyperlink display text and target URL; the payload is
//     a Hyperlink object, [MS-OSHARED] 2.3.7.1)

package oletext

import (
	"encoding/binary"
	"errors"
	"fmt"
	"io"
	"math"
	"strconv"
	"strings"
	"unicode/utf16"
)

// BIFF record types used here.
const (
	recFormula    = 0x0006
	recEOF        = 0x000A
	recHeader     = 0x0014
	recFooter     = 0x0015
	recLbl        = 0x0018
	recContinue   = 0x003C
	recBoundSheet = 0x0085
	recMulRk      = 0x00BD
	recRString    = 0x00D6
	recSST        = 0x00FC
	recLabelSst   = 0x00FD
	recLabel      = 0x0204
	recNumber     = 0x0203
	recString     = 0x0207
	recRK         = 0x027E
	recBOF        = 0x0809
	recFilePass   = 0x002F
	recTxO        = 0x01B6
	recHLink      = 0x01B8
	recSeriesText = 0x100D
)

// biffRecord is one BIFF record: a 2-byte type, 2-byte size and payload.
type biffRecord struct {
	typ  uint16
	data []byte
}

// extractXls extracts cell text from the Workbook (or BIFF5-era Book)
// stream by walking its BIFF records.
func extractXls(f *cfbFile) (string, error) {
	stream, err := f.openStream("Workbook")
	if err != nil {
		stream, err = f.openStream("Book") // BIFF5 era name
		if err != nil {
			return "", errors.New("Workbook/Book stream not found")
		}
	}

	recs := splitBiffRecords(stream)
	if len(recs) == 0 {
		return "", errors.New("no BIFF records found")
	}

	x := &xlsExtractor{curRow: -1, sheetIdx: -1, atLineStart: true}
	for i := 0; i < len(recs); i++ {
		r := recs[i]
		switch r.typ {
		case recFilePass:
			return "", errors.New("workbook is encrypted (FilePass)")
		case recBOF:
			x.onBOF(r.data)
		case recEOF:
			x.onEOF()
		case recBoundSheet:
			x.onBoundSheet(r.data)
		case recSST:
			// Hand the SST record together with its Continue records.
			segs := [][]byte{r.data}
			for i+1 < len(recs) && recs[i+1].typ == recContinue {
				i++
				segs = append(segs, recs[i].data)
			}
			x.onSST(segs)
		case recLabelSst:
			x.onLabelSst(r.data)
		case recLabel, recRString:
			x.onLabel(r.data)
		case recNumber:
			x.onNumber(r.data)
		case recRK:
			x.onRK(r.data)
		case recMulRk:
			x.onMulRk(r.data)
		case recFormula:
			x.onFormula(r.data)
		case recString:
			x.onString(r.data)
		case recTxO:
			// The shape/textbox/comment text follows in Continue records.
			var segs [][]byte
			for i+1 < len(recs) && recs[i+1].typ == recContinue {
				i++
				segs = append(segs, recs[i].data)
			}
			x.onTxO(r.data, segs)
		case recHeader, recFooter:
			x.onHeaderFooter(r.data)
		case recLbl:
			x.onLbl(r.data)
		case recSeriesText:
			x.onSeriesText(r.data)
		case recHLink:
			x.onHLink(r.data)
		}
	}
	out := x.out.String()
	if out != "" && !strings.HasSuffix(out, "\n") {
		out += "\n"
	}
	return out, nil
}

// splitBiffRecords slices a BIFF stream into records, stopping at the
// zero padding that can follow the final EOF record.
func splitBiffRecords(stream []byte) []biffRecord {
	var recs []biffRecord
	pos := 0
	for pos+4 <= len(stream) {
		typ := binary.LittleEndian.Uint16(stream[pos:])
		ln := int(binary.LittleEndian.Uint16(stream[pos+2:]))
		if typ == 0 { // padding after the final EOF
			break
		}
		if pos+4+ln > len(stream) {
			break
		}
		recs = append(recs, biffRecord{typ, stream[pos+4 : pos+4+ln]})
		pos += 4 + ln
	}
	return recs
}

// xlsExtractor accumulates extracted cell text while walking the record
// stream. Cells in the same row are joined with tabs; a new row starts a
// new line.
type xlsExtractor struct {
	out      strings.Builder
	sst      []string
	sheets   []string
	sheetIdx int
	curRow   int
	biff5    bool
	// bofDepth is the current BOF/EOF substream nesting depth. Top-level
	// sheet substreams sit at depth 1; an embedded chart's substream is
	// nested deeper and must not be mistaken for a new sheet.
	bofDepth int
	// atLineStart is true when the output buffer is empty or ends with a
	// newline, so the next cell or shape knows whether to start a new line.
	atLineStart bool
	// A formula whose cached result is a string: the next String record
	// belongs to this cell.
	pendingStringRow int
	hasPendingString bool
}

// onBOF handles a BOF record: it detects the BIFF version on the workbook
// globals substream and starts a new sheet section on worksheet substreams.
func (x *xlsExtractor) onBOF(d []byte) {
	x.bofDepth++
	if len(d) < 4 {
		return
	}
	vers := binary.LittleEndian.Uint16(d[0:])
	dt := binary.LittleEndian.Uint16(d[2:])
	if dt == 0x0005 { // workbook globals
		x.biff5 = vers != 0x0600
		return
	}
	if x.bofDepth != 1 {
		// A nested substream, e.g. an embedded chart inside a worksheet.
		// Its records belong to the enclosing sheet, so do not start a new
		// sheet section here.
		return
	}
	// Top-level sheet substreams (worksheet/chart/macro) appear in
	// BoundSheet8 order.
	x.sheetIdx++
	x.curRow = -1
	if dt == 0x0010 { // worksheet
		name := ""
		if x.sheetIdx < len(x.sheets) {
			name = x.sheets[x.sheetIdx]
		}
		x.newline()
		fmt.Fprintf(&x.out, "=== Sheet: %s ===\n", name)
		x.atLineStart = true
	}
}

// onEOF closes the current substream, unwinding the nesting tracked by
// onBOF so that the next top-level BOF is recognized as a new sheet.
func (x *xlsExtractor) onEOF() {
	if x.bofDepth > 0 {
		x.bofDepth--
	}
}

// onBoundSheet records a sheet name from a BoundSheet8 record:
// lbPlyPos(4) hsState/dt(2) stName(ShortXLUnicodeString).
func (x *xlsExtractor) onBoundSheet(d []byte) {
	if len(d) < 7 {
		return
	}
	if x.biff5 {
		cch := int(d[6])
		if 7+cch <= len(d) {
			x.sheets = append(x.sheets, latin1String(d[7:7+cch]))
		}
		return
	}
	r := newSegReader([][]byte{d[6:]})
	name, err := readXLString(r, 1)
	if err == nil {
		x.sheets = append(x.sheets, name)
	}
}

// onSST parses the shared string table, reading strings across the
// Continue record boundaries in segs.
func (x *xlsExtractor) onSST(segs [][]byte) {
	r := newSegReader(segs)
	if _, err := r.readU32(); err != nil { // cstTotal
		return
	}
	cstUnique, err := r.readU32()
	if err != nil {
		return
	}
	for i := uint32(0); i < cstUnique; i++ {
		s, err := readXLString(r, 2)
		if err != nil {
			break // keep the strings recovered so far
		}
		x.sst = append(x.sst, s)
	}
}

// onLabelSst emits a cell whose text is an index into the SST.
func (x *xlsExtractor) onLabelSst(d []byte) {
	if len(d) < 10 {
		return
	}
	row := int(binary.LittleEndian.Uint16(d[0:]))
	isst := binary.LittleEndian.Uint32(d[6:])
	if int(isst) < len(x.sst) {
		x.emitCell(row, x.sst[isst])
	}
}

// onLabel emits a cell with an inline string (Label / RString records).
func (x *xlsExtractor) onLabel(d []byte) {
	if len(d) < 8 {
		return
	}
	row := int(binary.LittleEndian.Uint16(d[0:]))
	if x.biff5 {
		// BIFF5: cch(2) followed by codepage-dependent bytes.
		cch := int(binary.LittleEndian.Uint16(d[6:]))
		if 8+cch <= len(d) {
			x.emitCell(row, latin1String(d[8:8+cch]))
		}
		return
	}
	r := newSegReader([][]byte{d[6:]})
	if s, err := readXLString(r, 2); err == nil {
		x.emitCell(row, s)
	}
}

// onNumber emits a cell holding an IEEE double.
func (x *xlsExtractor) onNumber(d []byte) {
	if len(d) < 14 {
		return
	}
	row := int(binary.LittleEndian.Uint16(d[0:]))
	v := math.Float64frombits(binary.LittleEndian.Uint64(d[6:]))
	x.emitCell(row, formatNum(v))
}

// onRK emits a cell holding an RK-encoded number.
func (x *xlsExtractor) onRK(d []byte) {
	if len(d) < 10 {
		return
	}
	row := int(binary.LittleEndian.Uint16(d[0:]))
	x.emitCell(row, formatNum(decodeRK(binary.LittleEndian.Uint32(d[6:]))))
}

// onMulRk emits the cells of a MulRk record:
// rw(2) colFirst(2) [ixfe(2) RK(4)]... colLast(2).
func (x *xlsExtractor) onMulRk(d []byte) {
	if len(d) < 12 {
		return
	}
	row := int(binary.LittleEndian.Uint16(d[0:]))
	n := (len(d) - 6) / 6
	for i := 0; i < n; i++ {
		v := decodeRK(binary.LittleEndian.Uint32(d[4+i*6+2:]))
		x.emitCell(row, formatNum(v))
	}
}

// onFormula emits a formula's cached value. Layout: cell(6) +
// CachedValue(8) + flags... If the last two bytes of CachedValue are
// 0xFFFF the value is non-numeric; a leading byte of 0x00 means a string
// result carried by the following String record ([MS-XLS] 2.5.133
// FormulaValue).
func (x *xlsExtractor) onFormula(d []byte) {
	if len(d) < 14 {
		return
	}
	row := int(binary.LittleEndian.Uint16(d[0:]))
	if binary.LittleEndian.Uint16(d[12:]) == 0xFFFF {
		if d[6] == 0x00 {
			x.pendingStringRow = row
			x.hasPendingString = true
		}
		return
	}
	v := math.Float64frombits(binary.LittleEndian.Uint64(d[6:]))
	x.emitCell(row, formatNum(v))
}

// onString emits the string result of the preceding Formula record.
func (x *xlsExtractor) onString(d []byte) {
	if !x.hasPendingString {
		return
	}
	x.hasPendingString = false
	if x.biff5 {
		if len(d) < 2 {
			return
		}
		cch := int(binary.LittleEndian.Uint16(d[0:]))
		if 2+cch <= len(d) {
			x.emitCell(x.pendingStringRow, latin1String(d[2:2+cch]))
		}
		return
	}
	r := newSegReader([][]byte{d})
	if s, err := readXLString(r, 2); err == nil {
		x.emitCell(x.pendingStringRow, s)
	}
}

// emitCell appends cell text: tab-separated within a row, newline when
// the row changes.
func (x *xlsExtractor) emitCell(row int, text string) {
	if text == "" {
		return
	}
	if row == x.curRow && !x.atLineStart {
		x.out.WriteByte('\t')
	} else {
		x.newline()
	}
	x.out.WriteString(text)
	x.atLineStart = false
	x.curRow = row
}

// newline starts a fresh output line unless the buffer is already at one.
func (x *xlsExtractor) newline() {
	if !x.atLineStart {
		x.out.WriteByte('\n')
		x.atLineStart = true
	}
}

// onTxO emits the text of a TxO record ([MS-XLS] 2.4.324): a shape, text
// box or cell comment. The fixed part gives cchText at offset 10; the
// characters live in the Continue records in segs, each of which begins
// with a fresh fHighByte flag byte ([MS-XLS] 2.5.293).
func (x *xlsExtractor) onTxO(d []byte, segs [][]byte) {
	if len(d) < 12 || len(segs) == 0 {
		return
	}
	cch := int(binary.LittleEndian.Uint16(d[10:]))
	if cch == 0 {
		return
	}
	r := newSegReader(segs)
	flags, err := r.readByte()
	if err != nil {
		return
	}
	s, _ := r.readChars(cch, flags&0x01 != 0)
	x.emitAux(s)
}

// emitAux writes a run of text that does not belong to the cell grid --
// shape/text-box/comment text, a header or footer, a defined name, chart
// text or a hyperlink -- on its own line so it does not run into the grid.
func (x *xlsExtractor) emitAux(text string) {
	text = cleanAuxText(text)
	if strings.TrimSpace(text) == "" {
		return
	}
	x.newline()
	x.out.WriteString(text)
	x.atLineStart = false
	x.newline()
	x.curRow = -1
}

// cleanAuxText keeps the printable content of non-cell text and turns the
// in-text line breaks Excel uses (CR / LF / VT) into newlines.
func cleanAuxText(s string) string {
	var sb strings.Builder
	sb.Grow(len(s))
	for _, r := range s {
		switch r {
		case '\r', '\n', 0x0B:
			sb.WriteByte('\n')
		case '\t':
			sb.WriteByte('\t')
		default:
			if r >= 0x20 {
				sb.WriteRune(r)
			}
		}
	}
	return strings.Trim(sb.String(), "\n")
}

// onHeaderFooter emits the text of a Header (0x0014) or Footer (0x0015)
// record ([MS-XLS] 2.4.137 / 2.4.136). When no header/footer is set the
// body is empty. The string carries page-setup format codes that
// stripHeaderFooterCodes removes, leaving the literal text.
func (x *xlsExtractor) onHeaderFooter(d []byte) {
	if len(d) == 0 {
		return
	}
	var s string
	if x.biff5 {
		// BIFF5: a byte-counted ANSI string (cch:1).
		cch := int(d[0])
		if 1+cch > len(d) {
			return
		}
		s = latin1String(d[1 : 1+cch])
	} else {
		r := newSegReader([][]byte{d})
		var err error
		if s, err = readXLString(r, 2); err != nil && s == "" {
			return
		}
	}
	x.emitAux(stripHeaderFooterCodes(s))
}

// stripHeaderFooterCodes removes the page-setup formatting codes from a
// header/footer string ([MS-XLS] 2.4.137): &L/&C/&R start the left/center/
// right sections (rendered here as spaces), &"font,style" and &<digits> set
// the font and size, && is a literal ampersand, &K<rrggbb> is a font color,
// and every other &<letter> is a field code (page number, date, ...) that
// carries no literal text.
func stripHeaderFooterCodes(s string) string {
	var sb strings.Builder
	rs := []rune(s)
	for i := 0; i < len(rs); i++ {
		if rs[i] != '&' {
			sb.WriteRune(rs[i])
			continue
		}
		if i+1 >= len(rs) {
			break
		}
		switch c := rs[i+1]; {
		case c == '&':
			sb.WriteRune('&')
			i++
		case c == '"': // &"font,style"
			i += 2
			for i < len(rs) && rs[i] != '"' {
				i++
			}
		case c >= '0' && c <= '9': // &<digits>: font size
			i++
			for i+1 < len(rs) && rs[i+1] >= '0' && rs[i+1] <= '9' {
				i++
			}
		case c == 'L' || c == 'C' || c == 'R': // section separators
			sb.WriteByte(' ')
			i++
		case c == 'K': // &K<rrggbb>: font color
			i++
			for j := 0; j < 6 && i+1 < len(rs) && isHexDigit(rs[i+1]); j++ {
				i++
			}
		default: // single-letter field code (&P, &D, &F, ...)
			i++
		}
	}
	return strings.TrimSpace(sb.String())
}

func isHexDigit(r rune) bool {
	return (r >= '0' && r <= '9') || (r >= 'a' && r <= 'f') || (r >= 'A' && r <= 'F')
}

// onLbl emits the name of a defined name (Lbl, [MS-XLS] 2.4.150). A 14-byte
// fixed part (grbit, chKey, cch, cce, ...) is followed by the name string.
// Built-in names (fBuiltin) are internal codes such as Print_Area and are
// skipped.
func (x *xlsExtractor) onLbl(d []byte) {
	if len(d) < 15 {
		return
	}
	grbit := binary.LittleEndian.Uint16(d[0:])
	if grbit&0x0020 != 0 { // fBuiltin
		return
	}
	cch := int(d[3])
	if cch == 0 {
		return
	}
	if x.biff5 {
		if 14+cch <= len(d) {
			x.emitAux(latin1String(d[14 : 14+cch]))
		}
		return
	}
	r := newSegReader([][]byte{d[14:]})
	flags, err := r.readByte()
	if err != nil {
		return
	}
	if s, err := r.readChars(cch, flags&0x01 != 0); err == nil {
		x.emitAux(s)
	}
}

// onSeriesText emits a chart text string (SeriesText, [MS-XLS] 2.4.252): a
// 2-byte unused id followed by a ShortXLUnicodeString (cch:1). These records
// hold chart titles, axis titles and series/category names.
func (x *xlsExtractor) onSeriesText(d []byte) {
	if len(d) < 4 {
		return
	}
	cch := int(d[2])
	if cch == 0 {
		return
	}
	r := newSegReader([][]byte{d[3:]})
	flags, err := r.readByte()
	if err != nil {
		return
	}
	if s, err := r.readChars(cch, flags&0x01 != 0); err == nil {
		x.emitAux(s)
	}
}

// Hyperlink flag bits ([MS-OSHARED] 2.3.7.1).
const (
	hlinkHasMoniker      = 0x00000001
	hlinkHasLocationStr  = 0x00000008
	hlinkHasDisplayName  = 0x00000010
	hlinkHasFrameName    = 0x00000080
	hlinkMonikerAsString = 0x00000100
)

// urlMonikerCLSID is the URL Moniker class id
// {79EAC9E0-BAF9-11CE-8C82-00AA004BA90B}, little-endian on the wire.
var urlMonikerCLSID = [16]byte{
	0xE0, 0xC9, 0xEA, 0x79, 0xF9, 0xBA, 0xCE, 0x11,
	0x8C, 0x82, 0x00, 0xAA, 0x00, 0x4B, 0xA9, 0x0B,
}

// onHLink emits the display text and target of a hyperlink (HLink,
// [MS-XLS] 2.4.323; the payload is a Hyperlink object, [MS-OSHARED]
// 2.3.7.1). Layout: ref8(8) + CLSID(16) + streamVersion(4) + flags(4),
// then the optional display name, target frame, moniker (the URL/path) and
// location strings selected by the flags.
func (x *xlsExtractor) onHLink(d []byte) {
	pos := 8 + 16 + 4 // ref8 + hlinkClsid + streamVersion
	if pos+4 > len(d) {
		return
	}
	flags := binary.LittleEndian.Uint32(d[pos:])
	pos += 4

	var target, loc string
	if flags&hlinkHasDisplayName != 0 {
		// The display text is the cell's own visible text and is already
		// emitted by the cell record, so it is read only to advance pos.
		_, pos = readHyperlinkString(d, pos)
	}
	if flags&hlinkHasFrameName != 0 {
		_, pos = readHyperlinkString(d, pos)
	}
	if flags&hlinkHasMoniker != 0 {
		if flags&hlinkMonikerAsString != 0 {
			target, pos = readHyperlinkString(d, pos)
		} else if pos+16 <= len(d) {
			isURL := isURLMoniker(d[pos : pos+16])
			pos += 16
			if isURL && pos+4 <= len(d) {
				// URLMoniker ([MS-OSHARED] 2.3.7.2): a 4-byte byte length
				// (including the terminating null) then a UTF-16LE URL.
				nb := int(binary.LittleEndian.Uint32(d[pos:]))
				pos += 4
				if nb >= 2 && pos+nb <= len(d) {
					target = strings.TrimRight(decodeUTF16(d[pos:pos+nb]), "\x00")
					pos += nb
				}
			}
		}
	}
	if flags&hlinkHasLocationStr != 0 {
		loc, _ = readHyperlinkString(d, pos)
	}
	if loc != "" {
		if target != "" {
			target += "#" + loc
		} else {
			target = loc
		}
	}
	x.emitAux(target)
}

// readHyperlinkString reads a HyperlinkString ([MS-OSHARED] 2.3.7.9): a
// 4-byte character count (including the terminating null) followed by that
// many UTF-16LE characters. It returns the string and the new offset.
func readHyperlinkString(d []byte, pos int) (string, int) {
	if pos+4 > len(d) {
		return "", pos
	}
	n := int(binary.LittleEndian.Uint32(d[pos:]))
	pos += 4
	if n <= 0 || pos+2*n > len(d) {
		return "", pos
	}
	s := strings.TrimRight(decodeUTF16(d[pos:pos+2*n]), "\x00")
	return s, pos + 2*n
}

func isURLMoniker(b []byte) bool {
	if len(b) < 16 {
		return false
	}
	for i := 0; i < 16; i++ {
		if b[i] != urlMonikerCLSID[i] {
			return false
		}
	}
	return true
}

// decodeUTF16 decodes a little-endian UTF-16 byte slice; a trailing odd
// byte, if any, is ignored.
func decodeUTF16(b []byte) string {
	u := make([]uint16, len(b)/2)
	for i := range u {
		u[i] = binary.LittleEndian.Uint16(b[i*2:])
	}
	return string(utf16.Decode(u))
}

// decodeRK decodes an RkNumber ([MS-XLS] 2.5.276).
func decodeRK(v uint32) float64 {
	var f float64
	if v&0x02 != 0 { // 30-bit signed integer
		f = float64(int32(v) >> 2)
	} else { // upper 30 bits of an IEEE double
		f = math.Float64frombits(uint64(v&0xFFFFFFFC) << 32)
	}
	if v&0x01 != 0 {
		f /= 100
	}
	return f
}

// formatNum renders a cell number the way %g would, with full precision.
func formatNum(v float64) string {
	return strconv.FormatFloat(v, 'g', -1, 64)
}

// ---- BIFF Unicode strings ----

// segReader reads across Continue record boundaries.
type segReader struct {
	segs [][]byte
	idx  int
	off  int
}

func newSegReader(segs [][]byte) *segReader { return &segReader{segs: segs} }

// advance moves to the next non-empty segment when the current one is
// exhausted. Used for fields that continue raw (rgRun, ExtRst, integers).
func (r *segReader) advance() bool {
	for r.idx < len(r.segs) && r.off >= len(r.segs[r.idx]) {
		r.idx++
		r.off = 0
	}
	return r.idx < len(r.segs)
}

func (r *segReader) readByte() (byte, error) {
	if !r.advance() {
		return 0, io.ErrUnexpectedEOF
	}
	b := r.segs[r.idx][r.off]
	r.off++
	return b, nil
}

func (r *segReader) readU16() (uint16, error) {
	lo, err := r.readByte()
	if err != nil {
		return 0, err
	}
	hi, err := r.readByte()
	if err != nil {
		return 0, err
	}
	return uint16(lo) | uint16(hi)<<8, nil
}

func (r *segReader) readU32() (uint32, error) {
	lo, err := r.readU16()
	if err != nil {
		return 0, err
	}
	hi, err := r.readU16()
	if err != nil {
		return 0, err
	}
	return uint32(lo) | uint32(hi)<<16, nil
}

func (r *segReader) skip(n int) error {
	for n > 0 {
		if !r.advance() {
			return io.ErrUnexpectedEOF
		}
		take := len(r.segs[r.idx]) - r.off
		if take > n {
			take = n
		}
		r.off += take
		n -= take
	}
	return nil
}

// readChars reads cch characters. [MS-XLS] 2.5.293: when the character
// array is split across a Continue record, the continued part starts
// with a fresh fHighByte flag byte.
func (r *segReader) readChars(cch int, high bool) (string, error) {
	u := make([]uint16, 0, cch)
	for cch > 0 {
		if r.idx < len(r.segs) && r.off >= len(r.segs[r.idx]) {
			// Crossing into a Continue record: re-read the option byte.
			if !r.advance() {
				break
			}
			b, err := r.readByte()
			if err != nil {
				break
			}
			high = b&0x01 != 0
		}
		if r.idx >= len(r.segs) {
			break
		}
		avail := len(r.segs[r.idx]) - r.off
		if high {
			take := cch
			if take > avail/2 {
				take = avail / 2
			}
			if take == 0 {
				return string(utf16.Decode(u)), io.ErrUnexpectedEOF
			}
			for i := 0; i < take; i++ {
				u = append(u, binary.LittleEndian.Uint16(r.segs[r.idx][r.off:]))
				r.off += 2
			}
			cch -= take
		} else {
			take := cch
			if take > avail {
				take = avail
			}
			if take == 0 {
				return string(utf16.Decode(u)), io.ErrUnexpectedEOF
			}
			for i := 0; i < take; i++ {
				u = append(u, uint16(r.segs[r.idx][r.off]))
				r.off++
			}
			cch -= take
		}
	}
	if cch > 0 {
		return string(utf16.Decode(u)), io.ErrUnexpectedEOF
	}
	return string(utf16.Decode(u)), nil
}

// readXLString parses an XLUnicodeString, XLUnicodeRichExtendedString or
// ShortXLUnicodeString. cchSize is the byte width of the character count
// field (1 or 2).
func readXLString(r *segReader, cchSize int) (string, error) {
	var cch int
	if cchSize == 1 {
		b, err := r.readByte()
		if err != nil {
			return "", err
		}
		cch = int(b)
	} else {
		v, err := r.readU16()
		if err != nil {
			return "", err
		}
		cch = int(v)
	}
	flags, err := r.readByte()
	if err != nil {
		return "", err
	}
	high := flags&0x01 != 0
	var cRun uint16
	var cbExt uint32
	if flags&0x08 != 0 { // fRichSt
		if cRun, err = r.readU16(); err != nil {
			return "", err
		}
	}
	if flags&0x04 != 0 { // fExtSt
		if cbExt, err = r.readU32(); err != nil {
			return "", err
		}
	}
	s, err := r.readChars(cch, high)
	if err != nil {
		return s, err
	}
	if err := r.skip(int(cRun) * 4); err != nil {
		return s, nil // the string is valid even if trailing formatting is missing
	}
	if err := r.skip(int(cbExt)); err != nil {
		return s, nil
	}
	return s, nil
}
