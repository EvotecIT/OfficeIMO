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

## AI Ingestion Pattern (With Citations)

```csharp
using OfficeIMO.Reader;
using System.Text;

var chunks = DocumentReader.ReadFolder(
    folderPath: @"C:\KnowledgeBase",
    folderOptions: new ReaderFolderOptions { Recurse = true, DeterministicOrder = true },
    options: new ReaderOptions { MaxChars = 4000 }).ToList();

var context = new StringBuilder();
foreach (var chunk in chunks) {
    var source = chunk.Location.Path ?? "unknown";
    var pointer = chunk.Location.Page.HasValue
        ? $"page {chunk.Location.Page.Value}"
        : chunk.Location.HeadingPath ?? $"block {chunk.Location.BlockIndex ?? 0}";

    context.AppendLine($"[source: {source} | {pointer}]");
    context.AppendLine(chunk.Markdown ?? chunk.Text);
    context.AppendLine();
}
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
- `DocumentReader.ReadFolder(...)` is best-effort: unreadable/corrupt/oversized files emit warning chunks and ingestion continues.
- The reader is best-effort and does not attempt OCR.
- Legacy binary formats (`.doc`, `.xls`, `.ppt`) are not supported.
