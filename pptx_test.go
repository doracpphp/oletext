package oletext

import "testing"

// pptx tests. Body, speaker notes, master/layout placeholders and document
// properties run against the real testdata/sample.pptx; the slide-comment
// case uses an in-memory package because python-pptx cannot author comments.

func TestPptxBody(t *testing.T) {
	extractFileWant(t, "testdata/sample.pptx",
		"Slide one shape text.",
		"二枚目のスライド。",
	)
}

func TestPptxNotes(t *testing.T) {
	extractFileWant(t, "testdata/sample.pptx", "Speaker notes for slide one.")
}

func TestPptxMaster(t *testing.T) {
	extractFileWant(t, "testdata/sample.pptx", "Click to edit Master title style")
}

func TestPptxProperties(t *testing.T) {
	extractFileWant(t, "testdata/sample.pptx",
		"Deck Property Title",
		"Deck Author",
	)
}

// TestPptxSlideComment covers the classic slide comment (text in <p:text>).
func TestPptxSlideComment(t *testing.T) {
	pres := `<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst></p:presentation>`
	rels := `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Target="slides/slide1.xml"/></Relationships>`
	slide := `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>Slide body.</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`
	comment := `<p:cmLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cm><p:text>A slide comment.</p:text></p:cm></p:cmLst>`
	extractWant(t, zipBytes(map[string]string{
		"ppt/presentation.xml":            pres,
		"ppt/_rels/presentation.xml.rels": rels,
		"ppt/slides/slide1.xml":           slide,
		"ppt/comments/comment1.xml":       comment,
	}), "Slide body.", "A slide comment.")
}
