package oletext

import (
	"bytes"
	"encoding/binary"
	"testing"
)

// Legacy PowerPoint (.ppt) tests against the LibreOffice-generated samples
// (testdata/gen.py shapes|big), plus hand-built record streams for the
// cases those samples cannot cover.

// TestPptBodyAndShape covers placeholder (title) text and a free text-box
// (shape).
func TestPptBodyAndShape(t *testing.T) {
	extractFileWant(t, "testdata/shapes.ppt",
		"Slide Title タイトル",
		"Shape in PowerPoint 図形テキスト",
	)
}

// TestPptLarge checks every slide's text in the large big.ppt is extracted.
func TestPptLarge(t *testing.T) {
	verifyBig(t, "testdata/big.ppt", 250, `スライド\d{6}`, nil)
}

// pptRec builds one [MS-PPT] record: recVer/recInstance, recType, recLen
// and the body.
func pptRec(verInst, recType uint16, body []byte) []byte {
	le := binary.LittleEndian
	b := le.AppendUint16(nil, verInst)
	b = le.AppendUint16(b, recType)
	b = le.AppendUint32(b, uint32(len(body)))
	return append(b, body...)
}

// pptSlide builds a container record holding one TextCharsAtom.
func pptSlide(text string) []byte {
	return pptRec(0x000F, 0x03EE, pptRec(0, rtTextCharsAtom, vbaUTF16(text)))
}

// pptEdit appends a PersistDirectoryAtom mapping the persist ids (from
// firstID on) to offsets, followed by the UserEditAtom of that save, and
// returns the UserEditAtom's offset.
func pptEdit(stream *bytes.Buffer, lastEdit uint32, firstID uint32, offsets ...uint32) uint32 {
	le := binary.LittleEndian
	dirOff := uint32(stream.Len())
	dir := le.AppendUint32(nil, firstID|uint32(len(offsets))<<20)
	for _, o := range offsets {
		dir = le.AppendUint32(dir, o)
	}
	stream.Write(pptRec(0, rtPersistDirectoryAtom, dir))

	editOff := uint32(stream.Len())
	edit := make([]byte, 28)
	le.PutUint32(edit[8:], lastEdit) // offsetLastEdit
	le.PutUint32(edit[12:], dirOff)  // offsetPersistDirectory
	stream.Write(pptRec(0, rtUserEditAtom, edit))
	return editOff
}

// pptCurrentUser builds a Current User stream pointing at the newest
// UserEditAtom.
func pptCurrentUser(headerToken, currentEdit uint32) []byte {
	le := binary.LittleEndian
	body := le.AppendUint32(nil, 0x14) // size
	body = le.AppendUint32(body, headerToken)
	body = le.AppendUint32(body, currentEdit)
	return pptRec(0, 0x0FF6, body)
}

// TestPptSkipsSupersededRecords models an incremental save: the second edit
// appends a new version of slide persist object 2 and repoints its persist
// directory entry. Only the live version's text may be extracted.
func TestPptSkipsSupersededRecords(t *testing.T) {
	var stream bytes.Buffer
	stream.Write(pptSlide("Kept text")) // persist id 1, at offset 0
	oldOff := uint32(stream.Len())
	stream.Write(pptSlide("Old text")) // persist id 2, first version
	edit1 := pptEdit(&stream, 0, 1, 0, oldOff)
	newOff := uint32(stream.Len())
	stream.Write(pptSlide("New text")) // persist id 2, second version
	edit2 := pptEdit(&stream, edit1, 2, newOff)

	got := extractWant(t, buildCFB(map[string][]byte{
		"PowerPoint Document": stream.Bytes(),
		"Current User":        pptCurrentUser(0xE391C05F, edit2),
	}), "Kept text", "New text")
	wantAbsent(t, got, "Old text")
}

// TestPptWithoutCurrentUser checks the whole stream is scanned when there is
// no persist directory to consult.
func TestPptWithoutCurrentUser(t *testing.T) {
	extractWant(t, buildCFB(map[string][]byte{
		"PowerPoint Document": pptSlide("Only text"),
	}), "Only text")
}

// TestPptEncrypted checks an encrypted presentation is rejected both when
// the Current User stream says so and when only the CryptSession10Container
// (itself a container record) is present.
func TestPptEncrypted(t *testing.T) {
	var stream bytes.Buffer
	stream.Write(pptSlide("ciphertext noise"))
	edit := pptEdit(&stream, 0, 1, 0)
	if _, err := Extract(buildCFB(map[string][]byte{
		"PowerPoint Document": stream.Bytes(),
		"Current User":        pptCurrentUser(pptEncryptedToken, edit),
	})); err == nil {
		t.Error("encrypted header token: expected an error, got nil")
	}

	crypt := append(pptSlide("ciphertext noise"), pptRec(0x000F, rtCryptSession10, make([]byte, 16))...)
	if _, err := Extract(buildCFB(map[string][]byte{"PowerPoint Document": crypt})); err == nil {
		t.Error("CryptSession10Container: expected an error, got nil")
	}
}
