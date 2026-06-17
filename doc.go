// doc.go extracts text from Word binary files (.doc).
//
// Referenced sections of [MS-DOC]:
//   - 2.4.1 Retrieving Text (the core algorithm)
//   - 2.5 The File Information Block (FIB); fcClx/lcbClx live in FibRgFcLcb97
//   - 2.8.35 Clx, 2.8.36 Pcdt, 2.8.37 PlcPcd
//   - 2.9.177 Pcd, 2.9.73 FcCompressed
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
	fibFcClx       = 0x01A2 // FibRgFcLcb97.fcClx
	fibLcbClx      = 0x01A6 // FibRgFcLcb97.lcbClx
)

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
			return decodePieceTable(clx[pos+5:pos+5+lcb], wd)
		default:
			return "", fmt.Errorf("unknown clxt 0x%02X in Clx", clx[pos])
		}
	}
	return "", errors.New("Pcdt not found in Clx")
}

// decodePieceTable walks the PlcPcd ([MS-DOC] 2.8.37): n+1 character
// positions (4 bytes each) followed by n Pcd structures (8 bytes each).
// Each Pcd's fc field is an FcCompressed ([MS-DOC] 2.9.73) whose bit 30
// is fCompressed: when set, the piece is 8-bit ANSI text at offset fc/2;
// otherwise it is UTF-16LE text at offset fc ([MS-DOC] 2.4.1 steps 5-6).
func decodePieceTable(plc, wd []byte) (string, error) {
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
		count := int(cpEnd - cpStart)
		pcd := plc[(n+1)*4+i*8:]
		fcRaw := binary.LittleEndian.Uint32(pcd[2:]) // Pcd: flags(2) fc(4) prm(2)
		compressed := fcRaw&0x40000000 != 0
		fc := int(fcRaw & 0x3FFFFFFF)

		if compressed {
			// [MS-DOC] 2.4.1 step 6: 8-bit ANSI text at fc/2.
			start := fc / 2
			if start+count > len(wd) {
				continue
			}
			for _, b := range wd[start : start+count] {
				sb.WriteRune(cp1252ToRune(b))
			}
		} else {
			// [MS-DOC] 2.4.1 step 5: UTF-16LE text at fc.
			if fc+2*count > len(wd) {
				continue
			}
			u := make([]uint16, count)
			for j := 0; j < count; j++ {
				u[j] = binary.LittleEndian.Uint16(wd[fc+2*j:])
			}
			sb.WriteString(string(utf16.Decode(u)))
		}
	}
	return cleanDocText(sb.String()), nil
}

// cleanDocText maps Word's in-text control characters to plain text and
// drops field codes (0x13 field begin / 0x14 separator / 0x15 end),
// keeping only the field result.
func cleanDocText(s string) string {
	var sb strings.Builder
	sb.Grow(len(s))
	inFieldCode := 0
	for _, r := range s {
		switch r {
		case 0x13: // field begin: skip until separator
			inFieldCode++
			continue
		case 0x14: // field separator: result text follows
			if inFieldCode > 0 {
				inFieldCode--
			}
			continue
		case 0x15: // field end
			continue
		}
		if inFieldCode > 0 {
			continue
		}
		switch r {
		case '\r', 0x0B, 0x0C: // paragraph mark, line break, page/section break
			sb.WriteByte('\n')
		case 0x07: // cell / row mark
			sb.WriteByte('\t')
		case 0x1E: // non-breaking hyphen
			sb.WriteByte('-')
		case 0x1F: // optional hyphen
		case 0x01, 0x02, 0x03, 0x04, 0x05, 0x08: // pictures, annotation refs, etc.
		default:
			if r >= 0x20 || r == '\t' || r == '\n' {
				sb.WriteRune(r)
			}
		}
	}
	return sb.String()
}
