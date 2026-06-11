// Package oletext extracts plain text from legacy Microsoft Office binary
// files (.doc, .xls, .ppt).
//
// The container format is parsed per [MS-CFB] and the document format is
// detected from the stream names inside the container rather than the file
// extension:
//
//	WordDocument        -> Word       ([MS-DOC])
//	Workbook / Book     -> Excel      ([MS-XLS])
//	PowerPoint Document -> PowerPoint ([MS-PPT])
//
// Text retrieval follows the published Microsoft open specifications:
// [MS-DOC] 2.4.1 (Retrieving Text) for Word, the BIFF8 record stream for
// Excel, and the TextCharsAtom/TextBytesAtom records for PowerPoint.
//
// Encrypted files and Word 95 or earlier formats are not supported.
package oletext

import (
	"fmt"
	"os"
	"strings"
)

// Extract parses data as an OLE2 compound file, detects the contained
// Office document type and returns its text encoded as UTF-8.
func Extract(data []byte) (string, error) {
	f, err := parseCFB(data)
	if err != nil {
		return "", err
	}
	switch {
	case f.hasStream("WordDocument"):
		return extractDoc(f)
	case f.hasStream("Workbook"), f.hasStream("Book"):
		return extractXls(f)
	case f.hasStream("PowerPoint Document"):
		return extractPpt(f)
	default:
		return "", fmt.Errorf("no supported document stream found (streams: %s)",
			strings.Join(f.streamNames(), ", "))
	}
}

// ExtractFile reads the named file and extracts its text with Extract.
func ExtractFile(path string) (string, error) {
	data, err := os.ReadFile(path)
	if err != nil {
		return "", err
	}
	return Extract(data)
}

// latin1String converts bytes to a string where each byte is the low byte
// of a UTF-16 code unit (Latin-1).
func latin1String(b []byte) string {
	var sb strings.Builder
	sb.Grow(len(b))
	for _, c := range b {
		sb.WriteRune(rune(c))
	}
	return sb.String()
}

// cp1252ToRune maps a Windows-1252 byte to a rune. The special character
// mapping for 8-bit ANSI text in [MS-DOC] 2.9.73 (FcCompressed) matches
// the CP-1252 0x80-0x9F range.
func cp1252ToRune(b byte) rune {
	if b < 0x80 || b > 0x9F {
		return rune(b)
	}
	return cp1252High[b-0x80]
}

var cp1252High = [32]rune{
	0x20AC, 0x0081, 0x201A, 0x0192, 0x201E, 0x2026, 0x2020, 0x2021,
	0x02C6, 0x2030, 0x0160, 0x2039, 0x0152, 0x008D, 0x017D, 0x008F,
	0x0090, 0x2018, 0x2019, 0x201C, 0x201D, 0x2022, 0x2013, 0x2014,
	0x02DC, 0x2122, 0x0161, 0x203A, 0x0153, 0x009D, 0x017E, 0x0178,
}
