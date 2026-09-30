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
//
// The stream is an append-only log: an incremental ("fast") save appends
// new versions of the changed top-level records and leaves the old ones in
// place. Which records are current is recorded by the persist object
// directory (2.3.2 UserEditAtom, 2.3.3 PersistDirectoryAtom), reached from
// the "Current User" stream (2.3.1 CurrentUserAtom). Only those live
// records are walked, so superseded and deleted text is not extracted.

package oletext

import (
	"encoding/binary"
	"errors"
	"sort"
	"strings"
)

const (
	rtUserEditAtom         = 0x0FF5
	rtTextCharsAtom        = 0x0FA0
	rtTextBytesAtom        = 0x0FA8
	rtPersistDirectoryAtom = 0x1772
	rtCryptSession10       = 0x2F14
)

// pptEncryptedToken is the CurrentUserAtom headerToken of an encrypted
// presentation ([MS-PPT] 2.3.1); an unencrypted one carries 0xE391C05F.
const pptEncryptedToken = 0xF3D1C4DF

// extractPpt extracts text from the "PowerPoint Document" stream by
// recursively walking the record tree of each live top-level record and
// collecting text atoms.
func extractPpt(f *cfbFile) (string, error) {
	stream, err := f.openStream("PowerPoint Document")
	if err != nil {
		return "", err
	}
	currentUser, _ := f.openStream("Current User")
	live, encrypted := pptLiveRecords(currentUser, stream)

	var parts []string
	for _, off := range live {
		// Each live offset is the header of one top-level record.
		recLen := int(binary.LittleEndian.Uint32(stream[off+4:]))
		end := len(stream)
		if recLen <= end-(off+8) {
			end = off + 8 + recLen
		}
		walkPptRecords(stream[off:end], 0, &parts, &encrypted)
	}
	if len(parts) == 0 && !encrypted {
		// No usable persist directory (or it led nowhere): fall back to
		// scanning the whole stream.
		walkPptRecords(stream, 0, &parts, &encrypted)
	}
	if encrypted {
		// The records of an encrypted presentation are ciphertext; whatever
		// happens to parse as a text atom is noise.
		return "", errors.New("presentation is encrypted")
	}
	if len(parts) == 0 {
		return "", nil
	}
	return strings.Join(parts, "\n") + "\n", nil
}

// pptLiveRecords returns the stream offsets of the current top-level
// records in ascending order, and whether the presentation is encrypted.
// It follows the chain of UserEditAtoms from the newest edit (named by the
// Current User stream) back to the first, merging their persist directories
// so that the newest offset of each persist object wins. It returns nil
// offsets when the structures are missing or damaged.
func pptLiveRecords(currentUser, stream []byte) (live []int, encrypted bool) {
	// CurrentUserAtom: rh(8) size(4) headerToken(4) offsetToCurrentEdit(4).
	if len(currentUser) < 20 {
		return nil, false
	}
	if binary.LittleEndian.Uint32(currentUser[12:]) == pptEncryptedToken {
		encrypted = true
	}
	persist := map[uint32]uint32{}
	off := uint64(binary.LittleEndian.Uint32(currentUser[16:]))
	for off != 0 {
		// UserEditAtom: rh(8) lastSlideIdRef(4) version(4) offsetLastEdit(4)
		// offsetPersistDirectory(4) docPersistIdRef(4) persistIdSeed(4)
		// lastView(2) unused(2) [encryptSessionPersistIdRef(4)].
		if off+36 > uint64(len(stream)) ||
			binary.LittleEndian.Uint16(stream[off+2:]) != rtUserEditAtom {
			return nil, encrypted
		}
		if binary.LittleEndian.Uint32(stream[off+4:]) >= 32 {
			encrypted = true // encryptSessionPersistIdRef is present
		}
		prev := uint64(binary.LittleEndian.Uint32(stream[off+16:]))
		dir := uint64(binary.LittleEndian.Uint32(stream[off+20:]))
		if !pptMergePersistDirectory(stream, dir, persist) {
			return nil, encrypted
		}
		if prev >= off { // edits are appended, so the chain must run backwards
			break
		}
		off = prev
	}
	seen := map[uint32]bool{}
	for _, o := range persist {
		if uint64(o)+8 <= uint64(len(stream)) && !seen[o] {
			seen[o] = true
			live = append(live, int(o))
		}
	}
	sort.Ints(live)
	return live, encrypted
}

// pptMergePersistDirectory reads the PersistDirectoryAtom at off and records
// the offset of every persist object not already known from a newer edit.
// Each PersistDirectoryEntry is a 20-bit first persist id and a 12-bit count
// packed into 4 bytes, followed by that many 4-byte stream offsets.
func pptMergePersistDirectory(stream []byte, off uint64, persist map[uint32]uint32) bool {
	if off+8 > uint64(len(stream)) ||
		binary.LittleEndian.Uint16(stream[off+2:]) != rtPersistDirectoryAtom {
		return false
	}
	body := stream[off+8:]
	if recLen := uint64(binary.LittleEndian.Uint32(stream[off+4:])); recLen < uint64(len(body)) {
		body = body[:recLen]
	}
	for len(body) >= 4 {
		v := binary.LittleEndian.Uint32(body)
		body = body[4:]
		id, n := v&0xFFFFF, v>>20
		for k := uint32(0); k < n && len(body) >= 4; k++ {
			if _, ok := persist[id+k]; !ok {
				persist[id+k] = binary.LittleEndian.Uint32(body)
			}
			body = body[4:]
		}
	}
	return true
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
		if recLen < 0 || recLen > len(data)-pos {
			recLen = len(data) - pos // treat a broken length as "rest of the data"
		}
		body := data[pos : pos+recLen]
		pos += recLen

		if recType == rtCryptSession10 && depth == 0 {
			// CryptSession10Container is a top-level record and itself a
			// container, so it has to be recognized before the generic
			// container descent.
			*encrypted = true
			continue
		}
		if verInst&0x000F == 0x000F { // container record
			walkPptRecords(body, depth+1, parts, encrypted)
			continue
		}
		switch recType {
		case rtTextCharsAtom:
			appendPptText(parts, decodeUTF16(body))
		case rtTextBytesAtom:
			// [MS-PPT] 2.9.45: each byte is the low byte of a UTF-16
			// character (high byte zero), i.e. Latin-1.
			appendPptText(parts, latin1String(body))
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
