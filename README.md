# oletext

Extract plain text from legacy Microsoft Office binary files
(.doc / .xls / .ppt). Pure Go, no external dependencies.

Based on the Microsoft open specifications:
[MS-CFB], [MS-DOC], [MS-XLS] (BIFF8) and [MS-PPT].

## Build

```
go build ./cmd/oletext
```

## Usage

```
oletext <file.doc|file.xls|file.ppt> ...
```

Extracted text is written to stdout as UTF-8. The document type is
detected from the streams inside the OLE2 container, not the file
extension.

## Library

```go
import "oletext"

text, err := oletext.ExtractFile("report.doc")
```

See `go doc oletext` for details.

## Limitations

- Encrypted files are rejected.
- Word 95 and earlier are not supported; BIFF5 (.xls) byte strings are
  decoded as Latin-1 (Word 97+/BIFF8/PPT 97+ store text as UTF-16, so
  non-Latin scripts work there).
- PPT output includes slide master and notes text.
