// ppt.go extracts text from PowerPoint binary files (.ppt).
//
// Referenced sections of [MS-PPT]:
//   - 2.3.1 RecordHeader (recVer/recInstance:2, recType:2, recLen:4;
//     recVer==0xF marks a container record)
//   - 2.9.46 TextCharsAtom (0x0FA0, UTF-16LE)
//   - 2.9.45 TextBytesAtom (0x0FA8, low bytes of UTF-16 characters)
//
// Shape text (text boxes, placeholders, table cells and grouped shapes)
// lives in the drawing layer: 2.5.10 PPDrawing -> [MS-ODRAW] OfficeArt
// drawing/group/shape containers (recType 0xF002/0xF003/0xF004) ->
// 2.9.76 OfficeArtClientTextbox (0xF00D) -> TextCharsAtom/TextBytesAtom.
// Every link in that chain is an OfficeArt container (recVer==0xF), so the
// recursive walk in walkPptRecords descends into it and collects the shape
// text the same way it collects ordinary slide text.

package oletext

import (
	"encoding/binary"
	"errors"
	"strings"
	"unicode/utf16"
)

const (
	rtTextCharsAtom  = 0x0FA0
	rtTextBytesAtom  = 0x0FA8
	rtCryptSession10 = 0x2F14
)

// extractPpt extracts text from the "PowerPoint Document" stream by
// recursively walking its record tree and collecting text atoms.
func extractPpt(f *cfbFile) (string, error) {
	stream, err := f.openStream("PowerPoint Document")
	if err != nil {
		return "", err
	}
	var parts []string
	encrypted := false
	walkPptRecords(stream, 0, &parts, &encrypted)
	if encrypted && len(parts) == 0 {
		return "", errors.New("presentation is encrypted (CryptSession10Container)")
	}
	if len(parts) == 0 {
		return "", nil
	}
	return strings.Join(parts, "\n") + "\n", nil
}

// walkPptRecords scans a record sequence, recursing into container
// records and collecting the text of TextCharsAtom and TextBytesAtom.
func walkPptRecords(data []byte, depth int, parts *[]string, encrypted *bool) {
	if depth > 32 {
		return
	}
	pos := 0
	for pos+8 <= len(data) {
		verInst := binary.LittleEndian.Uint16(data[pos:])
		recType := binary.LittleEndian.Uint16(data[pos+2:])
		recLen := int(binary.LittleEndian.Uint32(data[pos+4:]))
		pos += 8
		if recLen < 0 || pos+recLen > len(data) {
			recLen = len(data) - pos // treat a broken length as "rest of the data"
		}
		body := data[pos : pos+recLen]
		pos += recLen

		if verInst&0x000F == 0x000F { // container record
			walkPptRecords(body, depth+1, parts, encrypted)
			continue
		}
		switch recType {
		case rtTextCharsAtom:
			u := make([]uint16, len(body)/2)
			for i := range u {
				u[i] = binary.LittleEndian.Uint16(body[i*2:])
			}
			appendPptText(parts, string(utf16.Decode(u)))
		case rtTextBytesAtom:
			// [MS-PPT] 2.9.45: each byte is the low byte of a UTF-16
			// character (high byte zero), i.e. Latin-1.
			appendPptText(parts, latin1String(body))
		case rtCryptSession10:
			*encrypted = true
		}
	}
}

// appendPptText cleans a text atom and appends it unless it is blank.
func appendPptText(parts *[]string, s string) {
	s = cleanPptText(s)
	if strings.TrimSpace(s) != "" {
		*parts = append(*parts, s)
	}
}

// cleanPptText normalizes PPT line breaks: CR separates paragraphs and
// VT separates lines within a paragraph.
func cleanPptText(s string) string {
	var sb strings.Builder
	sb.Grow(len(s))
	for _, r := range s {
		switch r {
		case '\r', 0x0B:
			sb.WriteByte('\n')
		default:
			if r >= 0x20 || r == '\t' || r == '\n' {
				sb.WriteRune(r)
			}
		}
	}
	return sb.String()
}
