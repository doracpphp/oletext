// vba.go extracts VBA macro source from the macro storage shared by the
// legacy Word, Excel and PowerPoint binary formats.
//
// Referenced sections of [MS-OVBA]:
//   - 2.3.4.1 _VBA_PROJECT Stream (the project lives in a "VBA" storage)
//   - 2.3.4.2 dir Stream (compressed; lists the modules and, for each, its
//     MODULESTREAMNAME and the MODULEOFFSET where its source begins)
//   - 2.3.4.3 Module Streams (PerformanceCache then the CompressedContainer
//     of the source, starting at MODULEOFFSET)
//   - 2.4.1 Compression and Decompression (the run-length/copy-token codec)
//
// A project is found by its dir stream, wherever its "VBA" storage sits
// (Macros/VBA in Word, _VBA_PROJECT_CUR/VBA in Excel, the root VBA storage
// of a vbaProject.bin); its module streams are looked up in that same
// storage.

package oletext

import (
	"encoding/binary"
	"strings"
	"unicode/utf8"
)

// dir stream record ids ([MS-OVBA] 2.3.4.2) used here.
const (
	vbaIDProjectCodePage = 0x0003 // PROJECTCODEPAGE
	vbaIDProjectModules  = 0x000F // PROJECTMODULES (Count follows)
	vbaIDProjectCookie   = 0x0013 // PROJECTCOOKIE (follows PROJECTMODULES)
	vbaIDModuleName      = 0x0019 // MODULENAME (MBCS)
	vbaIDModuleNameUni   = 0x0047 // MODULENAMEUNICODE
	vbaIDStreamName      = 0x001A // MODULESTREAMNAME (MBCS)
	vbaIDStreamNameUni   = 0x0032 // MODULESTREAMNAME unicode (reserved record)
	vbaIDModuleOffset    = 0x0031 // MODULEOFFSET (TextOffset)
	vbaIDModuleTerm      = 0x002B // MODULE terminator (4 reserved zero bytes follow)
	vbaIDDirTerm         = 0x0010 // dir stream terminator
)

// vbaModule is one entry from the dir stream: where the module's source is.
type vbaModule struct {
	name       string // friendly module name (for the section header)
	streamName string // name of the CFB stream holding the module
	textOffset int    // byte offset of the CompressedContainer of the source
}

// extractVBA returns the source code of every VBA macro project in the
// compound file, or "" if there is none. Modules are emitted in dir order,
// each under a "=== VBA Module: <name> ===" header.
func extractVBA(f *cfbFile) string {
	var b strings.Builder
	for id, e := range f.dirs {
		if e.objType != objStream || e.name != "dir" {
			continue
		}
		dirRaw, err := f.readStream(id)
		if err != nil || len(dirRaw) == 0 {
			continue
		}
		dir := decompressVBA(dirRaw)
		codePage := vbaCodePage(dir)
		for _, m := range parseVBADir(dir) {
			stream, err := f.openStreamIn(e.parent, m.streamName)
			if err != nil || m.textOffset < 0 || m.textOffset > len(stream) {
				continue
			}
			src := normalizeVBA(decodeCodePage(decompressVBA(stream[m.textOffset:]), codePage))
			if strings.TrimSpace(src) == "" {
				continue
			}
			name := m.name
			if name == "" {
				name = m.streamName
			}
			b.WriteString("=== VBA Module: ")
			b.WriteString(name)
			b.WriteString(" ===\n")
			b.WriteString(src)
			if !strings.HasSuffix(src, "\n") {
				b.WriteByte('\n')
			}
		}
	}
	return b.String()
}

// vbaCodePage returns the PROJECTCODEPAGE of a decompressed dir stream: the
// code page the module source (and the MBCS names) are stored in. The record
// sits in the PROJECTINFORMATION section at the start of the stream, where
// every record before it has the plain Id+Size+Data layout. It returns 0
// when the record is missing.
func vbaCodePage(dir []byte) int {
	pos := 0
	for pos+6 <= len(dir) {
		id := binary.LittleEndian.Uint16(dir[pos:])
		size := int(binary.LittleEndian.Uint32(dir[pos+2:]))
		pos += 6
		if size < 0 || pos+size > len(dir) || id == vbaIDProjectModules {
			break
		}
		if id == vbaIDProjectCodePage && size >= 2 {
			return int(binary.LittleEndian.Uint16(dir[pos:]))
		}
		pos += size
	}
	return 0
}

// decodeCodePage converts text stored in a Windows code page to UTF-8.
// Text that is already valid UTF-8 (which covers plain ASCII and code page
// 65001) is returned unchanged. Otherwise Windows-1252 is decoded exactly;
// for any other code page only the ASCII range can be decoded without a
// conversion table, so each non-ASCII character becomes U+FFFD rather than
// leaking undecodable bytes into the UTF-8 output.
func decodeCodePage(b []byte, codePage int) string {
	if utf8.Valid(b) {
		return string(b)
	}
	var sb strings.Builder
	sb.Grow(len(b))
	switch codePage {
	case 0, 1252:
		for _, c := range b {
			sb.WriteRune(cp1252ToRune(c))
		}
	case 932, 936, 949, 950: // double-byte: a lead byte and its trail byte are one character
		for i := 0; i < len(b); i++ {
			c := b[i]
			switch {
			case c < 0x80:
				sb.WriteByte(c)
			case codePage == 932 && c >= 0xA1 && c <= 0xDF: // half-width katakana
				sb.WriteRune(0xFF61 + rune(c-0xA1))
			default:
				sb.WriteRune(utf8.RuneError)
				if i+1 < len(b) && b[i+1] >= 0x40 { // trail byte
					i++
				}
			}
		}
	default:
		for _, c := range b {
			if c < 0x80 {
				sb.WriteByte(c)
			} else {
				sb.WriteRune(utf8.RuneError)
			}
		}
	}
	return sb.String()
}

// parseVBADir reads the module table from a decompressed dir stream. It
// skips ahead to PROJECTMODULES because the preceding PROJECTINFORMATION
// and PROJECTREFERENCES sections contain records that do not follow the
// plain Id+Size+Data record layout; from PROJECTMODULES on, every record
// does. (The MODULE and dir terminators are an Id followed by 4 reserved
// zero bytes, which read as a record of size 0.)
func parseVBADir(dir []byte) []vbaModule {
	start := indexProjectModules(dir)
	if start < 0 {
		return nil
	}
	pos := start + 8 // PROJECTMODULES: Id(2)+Size(4)+Count(2)

	var mods []vbaModule
	var cur vbaModule
	for pos+6 <= len(dir) {
		id := binary.LittleEndian.Uint16(dir[pos:])
		size := int(binary.LittleEndian.Uint32(dir[pos+2:]))
		pos += 6
		if size < 0 || pos+size > len(dir) {
			break
		}
		data := dir[pos : pos+size]
		pos += size

		switch id {
		case vbaIDModuleName:
			if cur.name == "" {
				cur.name = latin1String(data)
			}
		case vbaIDModuleNameUni:
			if s := decodeUTF16(data); s != "" {
				cur.name = s
			}
		case vbaIDStreamName:
			if cur.streamName == "" {
				cur.streamName = latin1String(data)
			}
		case vbaIDStreamNameUni:
			if s := decodeUTF16(data); s != "" {
				cur.streamName = s
			}
		case vbaIDModuleOffset:
			if len(data) >= 4 {
				cur.textOffset = int(binary.LittleEndian.Uint32(data))
			}
		case vbaIDModuleTerm:
			if cur.streamName != "" {
				mods = append(mods, cur)
			}
			cur = vbaModule{}
		case vbaIDDirTerm:
			if cur.streamName != "" {
				mods = append(mods, cur)
			}
			return mods
		}
	}
	if cur.streamName != "" {
		mods = append(mods, cur)
	}
	return mods
}

// indexProjectModules finds the PROJECTMODULES record, identified by its id
// (0x000F) and fixed 2-byte size. The same six bytes can occur by chance in
// the reference records before it, so a match followed by the PROJECTCOOKIE
// record the spec puts next is preferred over a bare first match.
func indexProjectModules(dir []byte) int {
	first := -1
	for i := 0; i+6 <= len(dir); i++ {
		if binary.LittleEndian.Uint16(dir[i:]) != vbaIDProjectModules ||
			binary.LittleEndian.Uint32(dir[i+2:]) != 2 {
			continue
		}
		if i+14 <= len(dir) &&
			binary.LittleEndian.Uint16(dir[i+8:]) == vbaIDProjectCookie &&
			binary.LittleEndian.Uint32(dir[i+10:]) == 2 {
			return i
		}
		if first < 0 {
			first = i
		}
	}
	return first
}

// normalizeVBA converts the CRLF line endings of VBA source to '\n' and
// drops stray NULs.
func normalizeVBA(s string) string {
	s = strings.ReplaceAll(s, "\r\n", "\n")
	s = strings.ReplaceAll(s, "\r", "\n")
	s = strings.ReplaceAll(s, "\x00", "")
	return s
}

// maxVBAOutput caps a single decompression so a malformed CompressedContainer
// cannot be made to expand without bound.
const maxVBAOutput = 64 << 20

// decompressVBA decompresses a CompressedContainer ([MS-OVBA] 2.4.1.1): a
// SignatureByte (0x01) followed by CompressedChunks. It returns nil if the
// data is not a valid container. The codec is robust to truncation and bad
// offsets: it stops and returns what it has rather than panicking.
func decompressVBA(data []byte) []byte {
	if len(data) == 0 || data[0] != 0x01 {
		return nil
	}
	out := make([]byte, 0, len(data)*2)
	pos := 1
	for pos+2 <= len(data) && len(out) < maxVBAOutput {
		header := binary.LittleEndian.Uint16(data[pos:])
		pos += 2
		chunkLen := int(header&0x0FFF) + 1 // CompressedChunkSize = (bits)+3 total; data is that minus the 2-byte header => (bits)+1
		compressed := header&0x8000 != 0
		end := pos + chunkLen
		if end > len(data) {
			end = len(data)
		}
		chunkStart := len(out)

		if !compressed {
			out = append(out, data[pos:end]...)
			pos = end
			continue
		}
		for pos < end {
			flag := data[pos]
			pos++
			for bit := 0; bit < 8 && pos < end; bit++ {
				if flag&(1<<uint(bit)) == 0 {
					out = append(out, data[pos])
					pos++
					continue
				}
				if pos+2 > end {
					pos = end
					break
				}
				token := binary.LittleEndian.Uint16(data[pos:])
				pos += 2
				bitCount := copyTokenBits(len(out) - chunkStart)
				lengthMask := uint16(0xFFFF) >> bitCount
				length := int(token&lengthMask) + 3
				offset := int(token>>(16-bitCount)) + 1
				src := len(out) - offset
				if src < 0 {
					return out // malformed: give up on the rest
				}
				for k := 0; k < length && len(out) < maxVBAOutput; k++ {
					out = append(out, out[src+k])
				}
			}
		}
		pos = end
	}
	return out
}

// copyTokenBits is CopyTokenHelp ([MS-OVBA] 2.4.1.3.19.3): the number of bits
// the copy token spends on the offset, given how many bytes of the current
// chunk are already decompressed. It is ceil(log2(n)), clamped to [4, 12].
func copyTokenBits(decompressed int) uint {
	bits := uint(0)
	for (1 << bits) < decompressed {
		bits++
	}
	if bits < 4 {
		bits = 4
	}
	if bits > 12 {
		bits = 12
	}
	return bits
}
