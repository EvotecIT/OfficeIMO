# OfficeIMO.Reader

`OfficeIMO.Reader` is an optional, read-only facade that normalizes extraction across:
- Word (`.docx`, `.docm`) -> Markdown chunks
- Excel (`.xlsx`, `.xlsm`) -> table chunks + optional Markdown table previews
- PowerPoint (`.pptx`, `.pptm`) -> slide-aligned Markdown chunks (optionally including notes)
- Markdown (`.md`, `.markdown`) -> heading-aware text chunks
- PDF (`.pdf`) -> page-aware text chunks

The goal is to make it easy for tools like chat bots to ingest content deterministically.

## Quick Start

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
        MaxTotalBytes = 500L * 1024 * 1024,
        SkipReparsePoints = true,
        DeterministicOrder = true
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

## Notes
- `DocumentReader.Read(...)` is synchronous and streaming (returns `IEnumerable<T>`).
- `DocumentReader.ReadFolder(...)` is best-effort: unreadable/corrupt files are skipped so ingestion can continue.
- The reader is best-effort and does not attempt OCR.
- Legacy binary formats (`.doc`, `.xls`, `.ppt`) are not supported.
