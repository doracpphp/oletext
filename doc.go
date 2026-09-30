// doc.go extracts text from Word binary files (.doc).
//
// Referenced sections of [MS-DOC]:
//   - 2.4.1 Retrieving Text (the core algorithm)
//   - 2.5 The File Information Block (FIB); fcClx/lcbClx live in FibRgFcLcb97
//   - 2.8.35 Clx, 2.8.36 Pcdt, 2.8.37 PlcPcd
//   - 2.9.177 Pcd, 2.9.73 FcCompressed
//   - 2.8.10 PlcBtePapx, 2.9.175 PapxFkp, 2.9.174 PapxInFkp and 2.6 (Sprm)
//     for the paragraph properties that tell a table row mark from a cell mark
//
// A .doc holds text in more places than the main body. 2.3 (Document
// Parts) divides the document's character position (CP) range into
// subdocuments laid out one after another, each sized by a field of
// FibRgLw97 (2.5.4):
//
//	main document   (ccpText)     footnotes       (ccpFtn)
//	headers/footers (ccpHdd)      comments        (ccpAtn)
//	endnotes        (ccpEdn)      textboxes       (ccpTxbx)
//	header textboxes (ccpHdrTxbx)
//
// Shape text lives in the textbox subdocument and is referenced from the
// drawing layer via PlcftxbxTxt (2.8.21); comment text lives in the comment
// (annotation) subdocument; and so on. The piece table (PlcPcd) maps every
// CP in this whole range -- not just the CPs below ccpText -- to a file
// offset, so decodePieceTable, which walks all pieces, already reproduces
// every subdocument's text: body, footnotes, headers, comments, endnotes
// and textbox/shape text are all extracted.
//
// They are emitted concatenated in CP (subdocument) order with no per-part
// labels; the per-CP metadata that says which comment or footnote a run
// belongs to (PlcfandTxt, the reference Plcfs, etc.) is not consulted.

package oletext

import (
	"encoding/binary"
	"errors"
	"fmt"
	"strings"
	"unicode/utf16"
)

// FIB offsets within the WordDocument stream (fixed for Word 97 and later).
const (
	fibIdentOffset = 0x0000 // wIdent: 0xA5EC
	fibNFibOffset  = 0x0002
	fibFlagsOffset = 0x000A // fEncrypted=0x0100, fWhichTblStm=0x0200
	fibFcBtePapx   = 0x0102 // FibRgFcLcb97.fcPlcfBtePapx
	fibLcbBtePapx  = 0x0106 // FibRgFcLcb97.lcbPlcfBtePapx
	fibFcClx       = 0x01A2 // FibRgFcLcb97.fcClx
	fibLcbClx      = 0x01A6 // FibRgFcLcb97.lcbClx
)

// Sprms ([MS-DOC] 2.6.2, 2.6.3) inspected while looking for table row marks.
const (
	sprmPFTtp       = 0x2417 // the paragraph is a table row mark
	sprmPChgTabs    = 0xC615 // variable operand with an escape for long sizes
	sprmTDefTable   = 0xD608 // variable operand with a 2-byte size
	sprmPHugePapx   = 0x6646 // the real grpprl lives in the Data stream
	sprmPHugePapxV8 = 0x6645 // Word 97 variant of sprmPHugePapx
)

// docRowMark stands in for a table row mark in the decoded text. In the
// file both a cell mark and a row mark are U+0007; only the paragraph
// properties tell them apart, so decodePieceTable rewrites the row marks to
// this otherwise unused control character for cleanDocText.
const docRowMark = 0x1C

// extractDoc extracts the document text following [MS-DOC] 2.4.1:
// read the FIB, locate the Clx in the table stream via fcClx/lcbClx,
// then decode the piece table found in the Pcdt.
func extractDoc(f *cfbFile) (string, error) {
	wd, err := f.openStream("WordDocument")
	if err != nil {
		return "", err
	}
	if len(wd) < fibLcbClx+4 {
		return "", errors.New("WordDocument stream too short for a Word 97+ FIB")
	}
	if binary.LittleEndian.Uint16(wd[fibIdentOffset:]) != 0xA5EC {
		return "", errors.New("FIB signature mismatch (not a Word binary document)")
	}
	nFib := binary.LittleEndian.Uint16(wd[fibNFibOffset:])
	if nFib < 0x00C1 {
		return "", fmt.Errorf("nFib=0x%04X: Word 95 and earlier formats are not supported", nFib)
	}
	flags := binary.LittleEndian.Uint16(wd[fibFlagsOffset:])
	if flags&0x0100 != 0 {
		return "", errors.New("document is encrypted (fEncrypted=1)")
	}

	// [MS-DOC] 2.5.2: fWhichTblStm selects the 0Table or 1Table stream.
	tableName := "0Table"
	if flags&0x0200 != 0 {
		tableName = "1Table"
	}
	table, err := f.openStream(tableName)
	if err != nil {
		return "", fmt.Errorf("table stream: %w", err)
	}

	fcClx := binary.LittleEndian.Uint32(wd[fibFcClx:])
	lcbClx := binary.LittleEndian.Uint32(wd[fibLcbClx:])
	if lcbClx == 0 {
		return "", errors.New("Clx is empty")
	}
	if uint64(fcClx)+uint64(lcbClx) > uint64(len(table)) {
		return "", errors.New("Clx lies outside the table stream")
	}
	clx := table[fcClx : fcClx+lcbClx]

	// [MS-DOC] 2.8.35 Clx: a run of Prc structures (clxt=0x01) followed by
	// exactly one Pcdt (clxt=0x02).
	pos := 0
	for pos < len(clx) {
		switch clx[pos] {
		case 0x01: // Prc: clxt(1) + cbGrpprl(2) + GrpPrl
			if pos+3 > len(clx) {
				return "", errors.New("broken Prc in Clx")
			}
			cb := int(binary.LittleEndian.Uint16(clx[pos+1:]))
			pos += 3 + cb
		case 0x02: // Pcdt: clxt(1) + lcb(4) + PlcPcd
			if pos+5 > len(clx) {
				return "", errors.New("broken Pcdt in Clx")
			}
			lcb := int(binary.LittleEndian.Uint32(clx[pos+1:]))
			if pos+5+lcb > len(clx) {
				return "", errors.New("PlcPcd extends past Clx")
			}
			return decodePieceTable(clx[pos+5:pos+5+lcb], wd, docRowEnds(f, wd, table))
		default:
			return "", fmt.Errorf("unknown clxt 0x%02X in Clx", clx[pos])
		}
	}
	return "", errors.New("Pcdt not found in Clx")
}

// docRowEnds returns the WordDocument stream offsets at which a table row
// mark ends: the limit FC of every paragraph whose properties set fTtp.
// The paragraphs are enumerated through the PlcBtePapx, whose entries name
// the 512-byte PapxFkp pages holding each paragraph's FC range and property
// list. It returns nil when the structures are missing or damaged; the text
// is then extracted without row breaks.
func docRowEnds(f *cfbFile, wd, table []byte) map[uint32]bool {
	fc := binary.LittleEndian.Uint32(wd[fibFcBtePapx:])
	lcb := binary.LittleEndian.Uint32(wd[fibLcbBtePapx:])
	if lcb < 12 || uint64(fc)+uint64(lcb) > uint64(len(table)) {
		return nil
	}
	plc := table[fc : fc+lcb]
	n := (len(plc) - 4) / 8 // n+1 FCs followed by n page numbers

	// The Data stream is only needed for sprmPHugePapx; load it on demand.
	var data []byte
	dataLoaded := false
	loadData := func() []byte {
		if !dataLoaded {
			dataLoaded = true
			data, _ = f.openStream("Data")
		}
		return data
	}

	ends := map[uint32]bool{}
	for i := 0; i < n; i++ {
		pn := binary.LittleEndian.Uint32(plc[(n+1)*4+i*4:]) & 0x3FFFFF
		off := uint64(pn) * 512
		if off+512 > uint64(len(wd)) {
			continue
		}
		// PapxFkp: rgfc (cpara+1 FCs), rgbx (cpara 13-byte BxPap), the
		// PapxInFkp structures, and cpara in the last byte.
		page := wd[off : off+512]
		cpara := int(page[511])
		if (cpara+1)*4+cpara*13 > 511 {
			continue
		}
		for j := 0; j < cpara; j++ {
			bOffset := int(page[(cpara+1)*4+j*13]) * 2
			if bOffset == 0 {
				continue // no PapxInFkp: default properties
			}
			if grpprlHasTtp(papxGrpprl(page, bOffset), loadData) {
				ends[binary.LittleEndian.Uint32(page[(j+1)*4:])] = true
			}
		}
	}
	return ends
}

// papxGrpprl returns the property list of the PapxInFkp at offset off of an
// FKP page: a size byte cb (grpprlInPapx is 2*cb-1 bytes, or, when cb is 0,
// twice the following byte), then a GrpPrlAndIstd whose leading 2-byte istd
// is dropped.
func papxGrpprl(page []byte, off int) []byte {
	if off < 0 || off+2 > len(page) {
		return nil
	}
	size := 2*int(page[off]) - 1
	off++
	if size < 0 {
		size = 2 * int(page[off])
		off++
	}
	if size < 2 || off+size > len(page) {
		return nil
	}
	return page[off+2 : off+size]
}

// grpprlHasTtp reports whether a paragraph property list marks the
// paragraph as a table row mark (sprmPFTtp). The list is an array of Prl:
// a 2-byte Sprm whose top three bits (spra) give the operand size. data,
// when non-nil, supplies the Data stream for following a sprmPHugePapx to
// the property list stored there.
func grpprlHasTtp(g []byte, data func() []byte) bool {
	ttp := false
	for len(g) >= 2 {
		sprm := binary.LittleEndian.Uint16(g)
		g = g[2:]
		var n int
		switch sprm >> 13 {
		case 0, 1:
			n = 1
		case 2, 4, 5:
			n = 2
		case 3:
			n = 4
		case 7:
			n = 3
		case 6: // variable: a size byte, with two exceptions
			if len(g) < 2 {
				return ttp
			}
			n = 1 + int(g[0])
			if sprm == sprmTDefTable {
				n = 1 + int(binary.LittleEndian.Uint16(g))
			} else if sprm == sprmPChgTabs && g[0] == 255 {
				// cb(1) cTabsDel(1) rgdxaDel+rgdxaClose(4 each)
				// cTabsAdd(1) rgdxaAdd+rgtbdAdd(3 each).
				del := int(g[1])
				if 2+4*del >= len(g) {
					return ttp
				}
				n = 2 + 4*del + 1 + 3*int(g[2+4*del])
			}
		}
		if n > len(g) {
			return ttp
		}
		switch sprm {
		case sprmPFTtp:
			ttp = g[0] != 0
		case sprmPHugePapx, sprmPHugePapxV8:
			if data == nil {
				break
			}
			// The operand is the Data stream offset of a PrcData:
			// cbGrpprl(2) followed by the grpprl that replaces this one.
			d := data()
			off := uint64(binary.LittleEndian.Uint32(g))
			if off+2 > uint64(len(d)) {
				break
			}
			cb := uint64(binary.LittleEndian.Uint16(d[off:]))
			if off+2+cb > uint64(len(d)) {
				break
			}
			return grpprlHasTtp(d[off+2:off+2+cb], nil)
		}
		g = g[n:]
	}
	return ttp
}

// decodePieceTable walks the PlcPcd ([MS-DOC] 2.8.37): n+1 character
// positions (4 bytes each) followed by n Pcd structures (8 bytes each).
// Each Pcd's fc field is an FcCompressed ([MS-DOC] 2.9.73) whose bit 30
// is fCompressed: when set, the piece is 8-bit ANSI text at offset fc/2;
// otherwise it is UTF-16LE text at offset fc ([MS-DOC] 2.4.1 steps 5-6).
// A U+0007 that ends at one of the stream offsets in rowEnds is a table row
// mark and is rewritten to docRowMark.
func decodePieceTable(plc, wd []byte, rowEnds map[uint32]bool) (string, error) {
	// If the size is not 4 + a multiple of 12, truncate n and continue.
	n := (len(plc) - 4) / 12
	if n <= 0 {
		return "", errors.New("piece table has no pieces")
	}
	var sb strings.Builder
	for i := 0; i < n; i++ {
		cpStart := binary.LittleEndian.Uint32(plc[i*4:])
		cpEnd := binary.LittleEndian.Uint32(plc[(i+1)*4:])
		if cpEnd <= cpStart {
			continue
		}
		count := uint64(cpEnd - cpStart)
		pcd := plc[(n+1)*4+i*8:]
		fcRaw := binary.LittleEndian.Uint32(pcd[2:]) // Pcd: flags(2) fc(4) prm(2)
		compressed := fcRaw&0x40000000 != 0
		fc := uint64(fcRaw & 0x3FFFFFFF)

		if compressed {
			// [MS-DOC] 2.4.1 step 6: 8-bit ANSI text at fc/2.
			start := fc / 2
			if start+count > uint64(len(wd)) {
				continue
			}
			for j, b := range wd[start : start+count] {
				if b == 0x07 && rowEnds[uint32(start)+uint32(j)+1] {
					b = docRowMark
				}
				sb.WriteRune(cp1252ToRune(b))
			}
		} else {
			// [MS-DOC] 2.4.1 step 5: UTF-16LE text at fc.
			if fc+2*count > uint64(len(wd)) {
				continue
			}
			u := make([]uint16, count)
			for j := range u {
				off := fc + 2*uint64(j)
				u[j] = binary.LittleEndian.Uint16(wd[off:])
				if u[j] == 0x07 && rowEnds[uint32(off)+2] {
					u[j] = docRowMark
				}
			}
			sb.WriteString(string(utf16.Decode(u)))
		}
	}
	return cleanDocText(sb.String()), nil
}

// cleanDocText maps Word's in-text control characters to plain text and
// drops field codes (0x13 field begin / 0x14 separator / 0x15 end),
// keeping only the field result. Fields nest, and a field may have no
// separator at all (index entries such as XE or TC are all code), so each
// open field is tracked on a stack: an entry is true while that field's
// code portion is still open, and inCode counts such entries.
//
// Table cells are separated by tabs and rows by newlines. A cell mark's tab
// is held back until more text follows, so the mark of a row's last cell
// does not leave a trailing tab before the row's newline.
func cleanDocText(s string) string {
	var sb strings.Builder
	sb.Grow(len(s))
	var fields []bool
	inCode := 0
	pendingTab := false
	write := func(r rune) {
		if pendingTab {
			pendingTab = false
			sb.WriteByte('\t')
		}
		sb.WriteRune(r)
	}
	for _, r := range s {
		switch r {
		case 0x13: // field begin: code portion follows
			fields = append(fields, true)
			inCode++
			continue
		case 0x14: // field separator: result text follows
			if n := len(fields); n > 0 && fields[n-1] {
				fields[n-1] = false
				inCode--
			}
			continue
		case 0x15: // field end
			if n := len(fields); n > 0 {
				if fields[n-1] {
					inCode--
				}
				fields = fields[:n-1]
			}
			continue
		}
		if inCode > 0 {
			continue
		}
		switch r {
		case '\r', 0x0B, 0x0C: // paragraph mark, line break, page/section break
			write('\n')
		case 0x07: // cell mark
			if pendingTab {
				sb.WriteByte('\t') // the previous cell was empty
			}
			pendingTab = true
		case docRowMark: // row mark: ends the line, swallowing the last cell's tab
			pendingTab = false
			sb.WriteByte('\n')
		case 0x1E: // non-breaking hyphen
			write('-')
		case 0x1F: // optional hyphen
		case 0x01, 0x02, 0x03, 0x04, 0x05, 0x08: // pictures, annotation refs, etc.
		default:
			if r >= 0x20 || r == '\t' || r == '\n' {
				write(r)
			}
		}
	}
	if pendingTab {
		sb.WriteByte('\t')
	}
	return sb.String()
}
