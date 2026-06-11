// xls.go extracts text from Excel binary files (.xls).
//
// Referenced sections of [MS-XLS]:
//   - 2.1.4 Record (type:2, size:2, payload)
//   - 2.4.265 SST / 2.4.58 Continue (when a string spans a Continue record,
//     the continued part starts with a fresh fHighByte flag byte)
//   - 2.5.293 XLUnicodeRichExtendedString, 2.5.294 XLUnicodeString
//   - 2.4.149 LabelSst, 2.4.148 Label, 2.4.180 Number, 2.4.220 RK,
//     2.4.175 MulRk, 2.4.127 Formula, 2.4.268 String, 2.4.28 BoundSheet8

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

	x := &xlsExtractor{curRow: -1, sheetIdx: -1}
	for i := 0; i < len(recs); i++ {
		r := recs[i]
		switch r.typ {
		case recFilePass:
			return "", errors.New("workbook is encrypted (FilePass)")
		case recBOF:
			x.onBOF(r.data)
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
	// A formula whose cached result is a string: the next String record
	// belongs to this cell.
	pendingStringRow int
	hasPendingString bool
}

// onBOF handles a BOF record: it detects the BIFF version on the workbook
// globals substream and starts a new sheet section on worksheet substreams.
func (x *xlsExtractor) onBOF(d []byte) {
	if len(d) < 4 {
		return
	}
	vers := binary.LittleEndian.Uint16(d[0:])
	dt := binary.LittleEndian.Uint16(d[2:])
	if dt == 0x0005 { // workbook globals
		x.biff5 = vers != 0x0600
		return
	}
	// Substreams (worksheet/chart/macro) appear in BoundSheet8 order.
	x.sheetIdx++
	x.curRow = -1
	if dt == 0x0010 { // worksheet
		name := ""
		if x.sheetIdx < len(x.sheets) {
			name = x.sheets[x.sheetIdx]
		}
		if x.out.Len() > 0 {
			x.out.WriteByte('\n')
		}
		fmt.Fprintf(&x.out, "=== Sheet: %s ===\n", name)
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
	if x.curRow >= 0 {
		if row == x.curRow {
			x.out.WriteByte('\t')
		} else {
			x.out.WriteByte('\n')
		}
	}
	x.curRow = row
	x.out.WriteString(text)
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
