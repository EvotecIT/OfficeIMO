# OfficeIMO.Reader

`OfficeIMO.Reader` is a read-only facade that normalizes extraction across:
- Word (`.docx`, `.docm`, `.doc`) -> Markdown chunks
- Excel (`.xlsx`, `.xlsm`, `.xls`) -> row/table chunks (+ optional Markdown previews)
- PowerPoint (`.pptx`, `.pptm`, `.ppt`, `.pot`, `.pps`) -> slide-aligned Markdown chunks (optionally including notes)
- Markdown (`.md`, `.markdown`) -> heading-aware chunks
- PDF (`.pdf`) -> page-aware text chunks
- CSV/TSV, EPUB, HTML, standalone images, JSON, Jupyter notebooks, RTF, SRT/WebVTT subtitles, structured text, Visio, XML, YAML, and ZIP through modular adapter packages

It is designed for deterministic ingestion. Format-specific parsing stays in the owning OfficeIMO package, the facade remains thin, and optional adapters do not force unrelated dependencies into the core package.

Remote HTTP(S) retrieval is available only through the separate `OfficeIMO.Reader.Web` transport. It requires a caller-owned `HttpClient`, remains outside `OfficeIMO.Reader.All`, and passes bounded response bytes into the same configured local handlers.

## Install / Reference

Reference the `OfficeIMO.Reader` NuGet package or add a project reference. Install only the modular adapter packages required by the host.

The package README at `OfficeIMO.Reader/README.md` is the detailed API guide. The sections below summarize the stable host contracts and common ingestion patterns.

## Current Reader Surfaces

- `OfficeDocumentReader.Read(...)` returns `ReaderChunk` sequences for indexing and folder traversal.
- `ReadDocument(...)` returns the stable version 5 rich result with pages, blocks, tables, links, forms, assets, visuals, OCR candidates, metadata, and structured diagnostics.
- `OfficeDocumentReaderBuilder` freezes handlers, options, processor order, and concurrency into an isolated reader for services and concurrent hosts.
- Async file, stream, byte, and bounded multi-document APIs preserve cancellation, ordering, input limits, and caller-owned stream lifetime.
- Ordered processors provide opt-in normalization, artifact classification, link/table cleanup, and asset filtering.
- `ReadStructured(...)` emits bounded scalar records, sections, named tables, forms, and readiness diagnostics without an AI dependency.
- `ReadHierarchical(...)` creates token-bounded RAG leaves with overlap, exact source spans, and document/container/heading ancestry.
- `IOfficeOcrEngine` keeps OCR execution in a dependency-free core contract; process and Tesseract implementations are separate optional packages.
- `OfficeDocumentWebReader` adds explicit bounded HTTP(S) retrieval without adding network behavior to the core Reader, all-adapters preset, or global tool.

For format adapters, prefer instance-scoped registration:

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Pdf;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddPdfHandler()
    .WithMaxConcurrentReads(4)
    .AddProcessor(new OfficeDocumentBlockNormalizationProcessor())
    .Build();

OfficeDocumentReadResult document = await reader.ReadDocumentAsync(
    "report.pdf",
    cancellationToken: cancellationToken);
```

`OfficeDocumentReader.Default` is the built-in-only convenience instance. Build an isolated instance when modular handlers or processors are required.

## Basic Use (File Path)

```csharp
using OfficeIMO.Reader;

foreach (var chunk in OfficeDocumentReader.Default.Read(@"C:\Docs\Policy.docx")) {
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
var chunksFromStream = OfficeDocumentReader.Default.Read(fs, "Policy.docx").ToList();

// Bytes
var bytes = File.ReadAllBytes(@"C:\Docs\Policy.docx");
var chunksFromBytes = OfficeDocumentReader.Default.Read(bytes, "Policy.docx").ToList();
```

## Folders

```csharp
using OfficeIMO.Reader;

var chunks = OfficeDocumentReader.Default.ReadFolder(
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

var result = OfficeDocumentReader.Default.ReadFolderDetailed(
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

foreach (var doc in OfficeDocumentReader.Default.ReadFolderDocuments(
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

var chunks = OfficeDocumentReader.Default.ReadFolder(
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
    OpenPassword = "open-password",
    IncludeWordFootnotes = true,
    IncludePowerPointNotes = true,
    ExcelHeadersInFirstRow = true,
    ExcelChunkRows = 200,
    ExcelSheetName = "Data",
    ExcelA1Range = "A1:Z500",
    MarkdownChunkByHeadings = true,
    ComputeHashes = true
};

var chunks = OfficeDocumentReader.Default.Read(@"C:\Docs\Workbook.xlsx", options).ToList();
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
  - `BlockIndex`: emitted chunk index (0-based Reader order)
  - `SourceBlockIndex`: producer-defined index in the source document (when available)
  - `StartLine`: 1-based start line number (Markdown/text)
  - `HeadingPath`, `Sheet`, `A1Range`, `Slide`, `Page`
- `Warnings`: truncation/unsupported content warnings (best-effort).

## Stable Rich Result Transport

Use `ReadDocument(...)` when a host needs pages, blocks, tables, assets, links, forms, OCR candidates, visuals, or structured diagnostics in addition to chunks:

```csharp
OfficeDocumentReadResult document = OfficeDocumentReader.Default.ReadDocument(@"C:\Docs\Policy.docx");
string json = OfficeDocumentReadResultJson.Serialize(document);
OfficeDocumentReadResult restored = OfficeDocumentReadResultJson.Deserialize(json);
```

Schema version 5 is the first stable transport version. `OfficeDocumentReadResultSchema.GetJsonSchema()` returns the embedded JSON Schema, and the NuGet package includes the same artifact under `schemas/`. Versions 1 through 4 were experimental and are not accepted by the stable deserializer. Reader packages also validate their public .NET APIs against the current published NuGet versions during `dotnet pack`.

## Performance Evidence

`OfficeIMO.Reader.Benchmarks` measures rich extraction for every supported format family, bounded detection, JSON transport, and Markdown parser/chunker isolation using a deterministic generated corpus. Run the complete short suite with:

```powershell
dotnet run --project OfficeIMO.Reader.Benchmarks/OfficeIMO.Reader.Benchmarks.csproj -c Release -f net8.0
```

Keep raw BenchmarkDotNet artifacts local. When a result is used as a release baseline, record the hardware, runtime, corpus, and concise comparison in `Docs/benchmarks`.

## Notes / Limitations

- Legacy binary Word, Excel, and PowerPoint files (`.doc`, `.xls`, `.ppt`, `.pot`, `.pps`) route through their owning OfficeIMO import engines. Unsupported or preserve-only content is reported as reader warnings/diagnostics; projected PowerPoint images use the same asset contract as PPTX images.
- Folder ingestion is best-effort: unreadable/corrupt/oversized files emit warning chunks and processing continues.
- `ReadFolderDocuments(...)` yields per-source payloads (`ReaderSourceDocument`) for straightforward source/chunk table upserts.
- `ReadFolderDetailed(...)` provides aggregate counts and per-file status with optional progress callbacks.
- The core reader identifies OCR candidates and owns the bounded `IOfficeOcrEngine` execution/merge contract, but it does not include an OCR implementation. Use `OfficeIMO.Reader.Ocr.Process`, `OfficeIMO.Reader.Ocr.Tesseract`, or a host-supplied delegate engine when OCR should run.
