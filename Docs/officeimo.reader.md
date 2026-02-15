# OfficeIMO.Reader

`OfficeIMO.Reader` is an optional, read-only facade that normalizes extraction across:
- Word (`.docx`, `.docm`) -> Markdown chunks
- Excel (`.xlsx`, `.xlsm`) -> row/table chunks (+ optional Markdown previews)
- PowerPoint (`.pptx`, `.pptm`) -> slide-aligned Markdown chunks (optionally including notes)
- Markdown (`.md`, `.markdown`) -> heading-aware chunks
- PDF (`.pdf`) -> page-aware text chunks

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

## Folder Progress + Detailed Summary

```csharp
using OfficeIMO.Reader;

var result = DocumentReader.ReadFolderDetailed(
    folderPath: @"C:\KnowledgeBase",
    folderOptions: new ReaderFolderOptions { Recurse = true, MaxFiles = 10_000 },
    options: new ReaderOptions { ComputeHashes = true },
    includeChunks: true,
    onProgress: p => Console.WriteLine($"{p.Kind}: scanned={p.FilesScanned}, parsed={p.FilesParsed}, skipped={p.FilesSkipped}, chunks={p.ChunksProduced}"));

Console.WriteLine($"Files parsed: {result.FilesParsed}");
Console.WriteLine($"Files skipped: {result.FilesSkipped}");
Console.WriteLine($"Chunks: {result.ChunksProduced}");
```

## Database-Ready Folder Streaming

```csharp
using OfficeIMO.Reader;

foreach (var doc in DocumentReader.ReadFolderDocuments(
    folderPath: @"C:\KnowledgeBase",
    folderOptions: new ReaderFolderOptions { Recurse = true, MaxFiles = 10_000, DeterministicOrder = true },
    options: new ReaderOptions { ComputeHashes = true, MaxChars = 4_000 },
    onProgress: p => Console.WriteLine($"{p.Kind}: parsed={p.FilesParsed}, skipped={p.FilesSkipped}, chunks={p.ChunksProduced}"))) {

    if (!doc.Parsed) {
        Console.WriteLine($"SKIP {doc.Path}: {string.Join("; ", doc.Warnings ?? Array.Empty<string>())}");
        continue;
    }

    // Upsert source rows by SourceId/SourceHash, then upsert chunk rows by ChunkHash.
    Console.WriteLine($"{doc.Path} => {doc.ChunksProduced} chunks, ~{doc.TokenEstimateTotal} tokens");
}
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
    MarkdownChunkByHeadings = true,
    ComputeHashes = true
};

var chunks = DocumentReader.Read(@"C:\Docs\Workbook.xlsx", options).ToList();
```

## Output Contract

Each chunk is returned as `ReaderChunk`:
- `Id`: stable identifier (ASCII-only).
- `SourceId`: stable source-document identifier.
- `SourceHash`: optional source content hash.
- `ChunkHash`: optional per-chunk content hash.
- `TokenEstimate`: best-effort token estimate.
- `Kind`: input kind (Word/Excel/PowerPoint/Markdown/PDF/Text/Unknown).
- `Text`: plain text representation.
- `Markdown`: optional Markdown representation (when available).
- `Tables`: optional structured tables (Excel).
- `Location`: citation/debug metadata:
  - `Path`: source path or name (for citations)
  - `BlockIndex`: emitted chunk index (0-based order from `DocumentReader`)
  - `SourceBlockIndex`: producer-defined index in the source document (when available)
  - `StartLine`: 1-based start line number (Markdown/text)
  - `HeadingPath`, `Sheet`, `A1Range`, `Slide`, `Page`
- `Warnings`: truncation/unsupported content warnings (best-effort).

## Notes / Limitations

- Legacy binary formats (`.doc`, `.xls`, `.ppt`) are not supported (convert to OpenXML first).
- Folder ingestion is best-effort: unreadable/corrupt/oversized files emit warning chunks and processing continues.
- `ReadFolderDocuments(...)` yields per-source payloads (`ReaderSourceDocument`) for straightforward source/chunk table upserts.
- `ReadFolderDetailed(...)` provides aggregate counts and per-file status with optional progress callbacks.
- This reader does not do OCR.
