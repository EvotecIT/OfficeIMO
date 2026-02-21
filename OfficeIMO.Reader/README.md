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

    // Upsert your "sources" table keyed by doc.SourceId/doc.SourceHash,
    // then upsert chunk rows from doc.Chunks keyed by chunk.ChunkHash.
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

## Pluggable Handlers

`DocumentReader` now supports extension-based handler registration for modular packages.

```csharp
using OfficeIMO.Reader;

DocumentReader.RegisterHandler(new ReaderHandlerRegistration {
    Id = "sample.custom",
    DisplayName = "Sample Custom Reader",
    Kind = ReaderInputKind.Text,
    Extensions = new[] { ".sample" },
    ReadPath = (path, opts, ct) => new[] {
        new ReaderChunk {
            Id = "sample-0001",
            Kind = ReaderInputKind.Text,
            Location = new ReaderLocation { Path = path, BlockIndex = 0 },
            Text = "custom output"
        }
    }
});

var capabilities = DocumentReader.GetCapabilities();
```

Use `DocumentReader.UnregisterHandler("sample.custom")` to remove custom handlers.

`GetCapabilities()` exposes a stable contract surface for hosts:
- `SchemaId` = `officeimo.reader.capability`
- `SchemaVersion` = `1`
- stream/path support flags
- advertised warning behavior and deterministic output flag
- optional `DefaultMaxInputBytes` metadata for handler defaults

## Shared Input Guards For Adapter Authors

Use `ReaderInputLimits` when building modular handlers so `MaxInputBytes` behavior stays consistent across adapters:

```csharp
using OfficeIMO.Reader;

var parseStream = ReaderInputLimits.EnsureSeekableReadStream(
    stream,
    maxInputBytes: readerOptions?.MaxInputBytes,
    cancellationToken: ct,
    ownsStream: out var ownsParseStream);
```

You can also call `ReaderInputLimits.EnforceFileSize(path, maxBytes)` and `ReaderInputLimits.EnforceSeekableStreamSize(stream, maxBytes)` for path/seekable prechecks.

## Modular Adapter Registration (Optional Packages)

Keep dependencies split by registering only adapters you need:

```csharp
using OfficeIMO.Reader.Epub;
using OfficeIMO.Reader.Html;
using OfficeIMO.Reader.Text;
using OfficeIMO.Reader.Zip;

DocumentReaderEpubRegistrationExtensions.RegisterEpubHandler(replaceExisting: true);
DocumentReaderZipRegistrationExtensions.RegisterZipHandler(replaceExisting: true);
DocumentReaderHtmlRegistrationExtensions.RegisterHtmlHandler(replaceExisting: true);
DocumentReaderTextRegistrationExtensions.RegisterStructuredTextHandler(replaceExisting: true);
```

These adapters support both path and stream dispatch via `DocumentReader.Read(...)`.

## Notes
- `DocumentReader.Read(...)` is synchronous and streaming (returns `IEnumerable<T>`).
- `DocumentReader.ReadFolder(...)` is best-effort: unreadable/corrupt/oversized files emit warning chunks and ingestion continues.
- `DocumentReader.ReadFolderDocuments(...)` yields one source payload at a time (`ReaderSourceDocument`) for easy DB upserts.
- `DocumentReader.ReadFolderDetailed(...)` returns ingestion counts/file statuses and can surface progress callback events.
- Chunks include `SourceId`/`SourceHash`/`ChunkHash` + token estimate for incremental indexing and prompt budgeting.
- The reader is best-effort and does not attempt OCR.
- Legacy binary formats (`.doc`, `.xls`, `.ppt`) are not supported.
