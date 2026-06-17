package oletext

import (
	"strings"
	"testing"
)

// TestRejectsUnknownZip checks a ZIP that is not an OOXML package is
// reported as an error rather than silently returning nothing.
func TestRejectsUnknownZip(t *testing.T) {
	data := zipBytes(map[string]string{"random.txt": "just a zip"})
	if _, err := Extract(data); err == nil {
		t.Error("expected an error for a non-OOXML ZIP, got nil")
	}
}

// TestPptxSlideOrder checks slides are emitted in p:sldIdLst order, not part
// order: the list references rId2 (slide2) before rId1 (slide1).
func TestPptxSlideOrder(t *testing.T) {
	pres := `<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId id="256" r:id="rId2"/><p:sldId id="257" r:id="rId1"/></p:sldIdLst></p:presentation>`
	rels := `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Target="slides/slide1.xml"/><Relationship Id="rId2" Target="slides/slide2.xml"/></Relationships>`
	slide := func(text string) string {
		return `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>` + text + `</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`
	}
	got := extractWant(t, zipBytes(map[string]string{
		"ppt/presentation.xml":            pres,
		"ppt/_rels/presentation.xml.rels": rels,
		"ppt/slides/slide1.xml":           slide("First slide body"),
		"ppt/slides/slide2.xml":           slide("Second slide body"),
	}), "First slide body", "Second slide body")
	if strings.Index(got, "Second slide body") > strings.Index(got, "First slide body") {
		t.Errorf("slides not in presentation order:\n%s", got)
	}
}
