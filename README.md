# oletext

Extract plain text from Microsoft Office documents in pure Go, with no
external dependencies. Both the legacy OLE2 binary formats and the modern
Office Open XML (OOXML) formats are supported:

| Format | Extensions | Container | Spec |
|--------|------------|-----------|------|
| Word        | `.doc`  | OLE2 ([MS-CFB]) | [MS-DOC] |
| Excel       | `.xls`  | OLE2 ([MS-CFB]) | [MS-XLS] (BIFF8/BIFF5) |
| PowerPoint  | `.ppt`  | OLE2 ([MS-CFB]) | [MS-PPT] |
| Word        | `.docx` | ZIP             | OOXML ([ISO/IEC 29500]) |
| Excel       | `.xlsx` | ZIP             | OOXML |
| PowerPoint  | `.pptx` | ZIP             | OOXML |

The document type is detected from the file's **contents** (OLE2 stream
names / OOXML part names), not its extension.

## What is extracted

Beyond the main body text, oletext reaches the "secondary" text that
ordinary readers often miss:

- **Shapes / text boxes** (drawing-layer text in every format)
- **Comments / annotations** (cell comments, review comments, slide comments;
  classic and threaded)
- **Headers & footers** (with the `&L`/`&C`/`&R` and font format codes stripped)
- **Defined names**, **chart** titles/labels, and **hyperlink** targets (Excel)
- **Speaker notes**, **slide master / layout** placeholder text (PowerPoint)
- **Footnotes, endnotes, text boxes** and **document properties**
- **VBA macro source** ([MS-OVBA]) from the `Macros` / `_VBA_PROJECT_CUR`
  storage of legacy files and from the `vbaProject.bin` part of macro-enabled
  OOXML files

## Build

```
go build ./cmd/oletext
```

## CLI usage

```
oletext <file.doc|.xls|.ppt|.docx|.xlsx|.pptx> ...
```

Extracted text is written to stdout as UTF-8. With more than one argument,
each file's output is preceded by a `==> path <==` header.

## Library

```go
import "github.com/doracpphp/oletext"

// From a file on disk (OOXML packages are streamed from disk, so even very
// large .xlsx/.docx/.pptx are never loaded whole).
text, err := oletext.ExtractFile("report.docx")

// From bytes already in memory.
text, err := oletext.Extract(data)
```

`ExtractBytes` and `ExtractFileBytes` return the text as `[]byte`. See
`go doc github.com/doracpphp/oletext` for details.

## Limitations

- Encrypted documents are rejected.
- Word 95 and earlier are not supported. BIFF5 (`.xls`) byte strings are
  decoded as Latin-1; Word 97+/BIFF8/PPT 97+ and all OOXML formats store text
  as Unicode, so non-Latin scripts work there.
- Extraction is intentionally thorough: PowerPoint output can include slide
  master/layout placeholder prompts (e.g. "Click to edit Master title
  style"), and Word output concatenates all subdocuments (body, footnotes,
  comments, ...) without per-part labels.

## Tests and fixtures

Test fixtures are large and regenerable, so `testdata/` is git-ignored except
for `testdata/gen.py`, which builds every fixture:

```
# OOXML samples (needs: pip install python-docx openpyxl python-pptx)
py testdata/gen.py samples      # sample.docx/.xlsx/.pptx
py testdata/gen.py large        # big.docx/.xlsx/.pptx (perf)

# Legacy binary samples (needs LibreOffice; start a UNO listener first):
soffice --headless --invisible --nologo "--accept=socket,host=localhost,port=2003;urp;" &
"<LibreOffice>/program/python" testdata/gen.py shapes   # shapes.doc/.xls/.ppt
"<LibreOffice>/program/python" testdata/gen.py big        # big.doc/.xls/.ppt
```

Tests that need a fixture skip when it is absent, so `go test ./...` works on
a fresh checkout; run the generators above to exercise the real-file paths.

[MS-CFB]: https://learn.microsoft.com/openspecs/windows_protocols/ms-cfb/
[MS-DOC]: https://learn.microsoft.com/openspecs/office_file_formats/ms-doc/
[MS-XLS]: https://learn.microsoft.com/openspecs/office_file_formats/ms-xls/
[MS-PPT]: https://learn.microsoft.com/openspecs/office_file_formats/ms-ppt/
[MS-OVBA]: https://learn.microsoft.com/openspecs/office_file_formats/ms-ovba/
[ISO/IEC 29500]: https://www.iso.org/standard/71691.html
