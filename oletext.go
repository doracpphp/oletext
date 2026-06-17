// Package oletext extracts plain text from Microsoft Office documents,
// both the legacy binary formats (.doc, .xls, .ppt) and the post-2007
// Office Open XML formats (.docx, .xlsx, .pptx).
//
// The format is detected from the file contents, not its extension. A
// legacy document is an OLE2 compound file ([MS-CFB]) whose stream names
// identify the application:
//
//	WordDocument        -> Word       ([MS-DOC])
//	Workbook / Book     -> Excel      ([MS-XLS])
//	PowerPoint Document -> PowerPoint ([MS-PPT])
//
// Text retrieval follows the published Microsoft open specifications:
// [MS-DOC] 2.4.1 (Retrieving Text) for Word, the BIFF8 record stream for
// Excel, and the TextCharsAtom/TextBytesAtom records for PowerPoint.
//
// An OOXML document is instead a ZIP package of XML parts ([ISO/IEC
// 29500]) detected from the part names word/document.xml, xl/workbook.xml
// and ppt/presentation.xml; see ooxml.go.
//
// Encrypted files and Word 95 or earlier formats are not supported.
package oletext

import (
	"archive/zip"
	"fmt"
	"os"
	"strings"
	"unsafe"
)

// Extract detects the Office document type of data and returns its text
// encoded as UTF-8. The post-2007 formats (.docx, .xlsx, .pptx) are ZIP
// packages handled by extractOOXML; the legacy formats (.doc, .xls, .ppt)
// are OLE2 compound files detected from their stream names.
func Extract(data []byte) (string, error) {
	if isZip(data) {
		return extractOOXML(data)
	}
	f, err := parseCFB(data)
	if err != nil {
		return "", err
	}
	var text string
	switch {
	case f.hasStream("WordDocument"):
		text, err = extractDoc(f)
	case f.hasStream("Workbook"), f.hasStream("Book"):
		text, err = extractXls(f)
	case f.hasStream("PowerPoint Document"):
		text, err = extractPpt(f)
	default:
		return "", fmt.Errorf("no supported document stream found (streams: %s)",
			strings.Join(f.streamNames(), ", "))
	}
	if err != nil {
		return "", err
	}
	// VBA macro source lives in a separate storage shared by all three
	// formats; append it after the document body if present.
	if vba := extractVBA(f); vba != "" {
		if text != "" && !strings.HasSuffix(text, "\n") {
			text += "\n"
		}
		text += vba
	}
	return text, nil
}

// ExtractBytes is like [Extract] but returns the text as a UTF-8 byte
// slice, for callers that work with []byte. The slice aliases the freshly
// built result without copying it (see [bytesNoCopy]); treat it as
// read-only. It returns nil on error.
func ExtractBytes(data []byte) ([]byte, error) {
	s, err := Extract(data)
	if err != nil {
		return nil, err
	}
	return bytesNoCopy(s), nil
}

// ExtractFile extracts text from the named file. An OOXML document
// (.docx/.xlsx/.pptx) is read with zip.OpenReader, which streams the
// central directory and each part from disk on demand, so even a very
// large package is never loaded into memory whole. A legacy compound file
// (.doc/.xls/.ppt) requires random access and is read fully before being
// handed to Extract.
func ExtractFile(path string) (string, error) {
	if zr, err := zip.OpenReader(path); err == nil {
		defer zr.Close()
		return extractOOXMLReader(&zr.Reader)
	}
	data, err := os.ReadFile(path)
	if err != nil {
		return "", err
	}
	return Extract(data)
}

// ExtractFileBytes is like [ExtractFile] but returns the text as a UTF-8
// byte slice aliasing the result without copying (see [bytesNoCopy]); treat
// it as read-only. It returns nil on error.
func ExtractFileBytes(path string) ([]byte, error) {
	s, err := ExtractFile(path)
	if err != nil {
		return nil, err
	}
	return bytesNoCopy(s), nil
}

// bytesNoCopy returns a []byte that shares s's backing array instead of
// copying it. Extract and ExtractFile return freshly built, uniquely owned
// strings (from a strings.Builder or strings.Join), so no other string
// aliases these bytes and exposing them as a slice is sound -- on the
// condition the caller does not mutate the slice (a string's contents must
// stay immutable). This avoids copying the whole extracted text, which for
// a large spreadsheet is several megabytes.
func bytesNoCopy(s string) []byte {
	if len(s) == 0 {
		return nil
	}
	return unsafe.Slice(unsafe.StringData(s), len(s))
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
