# OfficeIMO.Reader

`OfficeIMO.Reader` is an optional, read-only facade that normalizes extraction across:
- Word (`.docx`, `.docm`) -> Markdown chunks
- Excel (`.xlsx`, `.xlsm`) -> row/table chunks (+ optional Markdown previews)
- PowerPoint (`.pptx`, `.pptm`) -> slide-aligned Markdown chunks (optionally including notes)
- Markdown (`.md`, `.markdown`) -> heading-aware chunks

It is designed to be deterministic and dependency-free (beyond OfficeIMO itself), so consumers (like IntelligenceX) can ingest content reliably.

## Install / Reference

Reference the `OfficeIMO.Reader` NuGet package (once published) or add a project reference.

## Basic Use (File Path)

```csharp
using OfficeIMO.Reader;

foreach (var chunk in DocumentReader.Read(@"C:\Docs\Policy.docx")) {
    Console.WriteLine(chunk.Id);
    Console.WriteLine(chunk.Location.HeadingPath);
    Console.WriteLine(chunk.Markdown ?? chunk.Text);
}
```

## Streams / Bytes

```csharp
using OfficeIMO.Reader;

// Stream (does not close the stream)
using var fs = File.OpenRead(@"C:\Docs\Policy.docx");
var chunksFromStream = DocumentReader.Read(fs, "Policy.docx").ToList();

// Bytes
var bytes = File.ReadAllBytes(@"C:\Docs\Policy.docx");
var chunksFromBytes = DocumentReader.Read(bytes, "Policy.docx").ToList();
```

## Folders

```csharp
using OfficeIMO.Reader;

var chunks = DocumentReader.ReadFolder(
    folderPath: @"C:\Docs",
    folderOptions: new ReaderFolderOptions {
        Recurse = true,
        MaxFiles = 500,
        MaxTotalBytes = 500L * 1024 * 1024
    },
    options: new ReaderOptions {
        MaxChars = 8_000
    }).ToList();
```

## Options

```csharp
using OfficeIMO.Reader;

var options = new ReaderOptions {
    MaxChars = 8_000,
    MaxTableRows = 200,
    IncludeWordFootnotes = true,
    IncludePowerPointNotes = true,
    ExcelHeadersInFirstRow = true,
    ExcelChunkRows = 200,
    ExcelSheetName = "Data",
    ExcelA1Range = "A1:Z500",
    MarkdownChunkByHeadings = true
};

var chunks = DocumentReader.Read(@"C:\Docs\Workbook.xlsx", options).ToList();
```

## Output Contract

Each chunk is returned as `ReaderChunk`:
- `Id`: stable identifier (ASCII-only).
- `Kind`: input kind (Word/Excel/PowerPoint/Markdown/Text/Unknown).
- `Text`: plain text representation.
- `Markdown`: optional Markdown representation (when available).
- `Tables`: optional structured tables (Excel).
- `Location`: citation/debug metadata:
  - `Path`: source path or name (for citations)
  - `BlockIndex`: emitted chunk index (0-based order from `DocumentReader`)
  - `SourceBlockIndex`: producer-defined index in the source document (when available)
  - `StartLine`: 1-based start line number (Markdown/text)
  - `HeadingPath`, `Sheet`, `A1Range`, `Slide`
- `Warnings`: truncation/unsupported content warnings (best-effort).

## Notes / Limitations

- Legacy binary formats (`.doc`, `.xls`, `.ppt`) are not supported (convert to OpenXML first).
- This reader is best-effort and does not do OCR.
