---
title: Reader and Extraction
description: Deterministic document extraction, rich results, structured records, RAG chunking, and optional OCR with OfficeIMO.Reader.
order: 55
---

`OfficeIMO.Reader` is the shared read-only extraction facade for Office and adjacent document formats. It can return simple `ReaderChunk` sequences for indexing, or a rich `OfficeDocumentReadResult` with pages, blocks, tables, links, forms, assets, visuals, OCR candidates, metadata, and structured diagnostics.

Format-specific parsing stays in the owning OfficeIMO package. Optional adapters and OCR providers remain separate packages, so applications install only the formats and runtimes they need.

## Install

Install the selective adapters used by the application. Each adapter brings the neutral Core contracts. Use `All`
only for a host that deliberately wants every adapter:

```powershell
dotnet add package OfficeIMO.Reader.Word
dotnet add package OfficeIMO.Reader.Pdf
dotnet add package OfficeIMO.Reader.Html
dotnet add package OfficeIMO.Reader.Epub

dotnet add package OfficeIMO.Reader.All
```

`OfficeIMO.Reader.Core` owns contracts, routing, plain text, processing, and custom-handler registration; it has no
format-engine dependencies. Word, Excel, PowerPoint, Markdown, Email, PDF, and the other formats are separate adapters
over their owning OfficeIMO engines.

## Supported formats

Each row is a separate package with its own `Add...Handler()` registration. `AddAllOfficeIMOHandlers()` from `OfficeIMO.Reader.All` registers every row marked **All**, in one isolated builder:

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.All;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddAllOfficeIMOHandlers(new ReaderAllOptions {
        IncludeLegacyWord = false,
        IncludeLegacySpreadsheet = false
    })
    .Build();
```

| Format family | Package | Extensions | Registration | All |
|---|---|---|---|---|
| Plain text | `OfficeIMO.Reader.Core` | `.txt`, `.log`, plus a last-resort handler for detected unknown payloads | `AddPlainTextHandlers()` | All |
| Word | `OfficeIMO.Reader.Word` | `.docx`, `.docm`, `.dotx`, `.dotm`, `.doc`, `.dot` | `AddWordHandler()` | All |
| Legacy word processors | `OfficeIMO.Reader.Word` | `.wp`, `.wp5`, `.wp6`, `.wpd`, `.ws`, `.ws3`-`.ws7`, `.sam`, `.lwp`, `.wps`, `.wri` | `AddLegacyWordHandler()` | All (`IncludeLegacyWord`, on by default) |
| Excel | `OfficeIMO.Reader.Excel` | `.xlsx`, `.xlsm`, `.xltx`, `.xltm`, `.xlam`, `.xlsb`, `.xls`, `.xlt`, `.xla`, `.xlm`, `.xlw` | `AddExcelHandler()` | All |
| Legacy spreadsheets | `OfficeIMO.Reader.Excel` | `.wk1`-`.wk4`, `.123`, `.wq1`, `.wq2`, `.wb1`-`.wb3`, `.qpw`, `.mp`, `.mp1`-`.mp3`, `.wks`, `.xlr`, `.slk`, `.dif` | `AddLegacySpreadsheetHandler()` | All (`IncludeLegacySpreadsheet`, on by default) |
| PowerPoint | `OfficeIMO.Reader.PowerPoint` | `.pptx`, `.pptm`, `.potx`, `.potm`, `.ppsx`, `.ppsm`, `.ppam`, `.ppt`, `.pot`, `.pps`, `.ppa` | `AddPowerPointHandler()` | All |
| PDF | `OfficeIMO.Reader.Pdf` | `.pdf` | `AddPdfHandler()` | All |
| OpenDocument | `OfficeIMO.Reader.OpenDocument` | `.odt`, `.ods`, `.odp`, `.odg`, `.fodt`, `.fods`, `.fodp`, `.fodg` | `AddOpenDocumentHandler()` | All |
| Apple iWork | `OfficeIMO.Reader.IWork` | `.pages`, `.numbers`, `.key` | `AddIWorkHandler()` | All |
| OneNote | `OfficeIMO.Reader.OneNote` | `.one`, `.onetoc2`, `.onepkg` | `AddOneNoteHandler()` | All (sibling-section reads from a `.onetoc2` are off by default) |
| Visio | `OfficeIMO.Reader.Visio` | `.vsdx`, `.vstx`, `.vssx`, `.vsdm`, `.vstm`, `.vssm`, `.vdx`, `.vsx`, `.vtx`, `.vsd`, `.vss`, `.vst` | `AddVisioHandler()` | All |
| Publisher | `OfficeIMO.Reader.Publisher` | `.pub` | `AddPublisherHandler()` | All |
| XPS | `OfficeIMO.Reader.Xps` | `.xps`, `.oxps` | `AddXpsHandler()` | All |
| DjVu | `OfficeIMO.Reader.DjVu` | `.djvu`, `.djv` | `AddDjVuHandler()` | All |
| Email messages | `OfficeIMO.Reader.Email` | `.eml`, `.msg`, `.oft`, `.tnef`, `.mbox`, `.mbx` | `AddEmailHandlers()` | All |
| Calendars and contacts | `OfficeIMO.Reader.Email` | `.ics`, `.ical`, `.ifb`, `.vcs`, `.vcf`, `.vcard` | `AddEmailHandlers()` | All |
| Mailbox stores | `OfficeIMO.Reader.Email` | `.pst`, `.ost`, `.olm`, `.emlx`, `.dbx`, `.oab` | `AddEmailHandlers()` | All |
| MHTML | `OfficeIMO.Reader.Email` | `.mht`, `.mhtml` | `AddEmailHandlers()` or `AddMhtmlHandler()` | All |
| HTML | `OfficeIMO.Reader.Html` | `.html`, `.htm`, `.xhtml` | `AddHtmlHandler()` | All |
| EPUB | `OfficeIMO.Reader.Epub` | `.epub` | `AddEpubHandler()` | All |
| Compiled HTML Help | `OfficeIMO.Reader.Chm` | `.chm` | `AddChmHandler()` | All |
| RTF | `OfficeIMO.Reader.Rtf` | `.rtf` | `AddRtfHandler()` | All |
| Markdown | `OfficeIMO.Reader.Markdown` | `.md`, `.markdown`, `.mdown`, `.mkd`, `.mdx` | `AddMarkdownHandler()` | All |
| AsciiDoc | `OfficeIMO.Reader.AsciiDoc` | `.adoc`, `.asciidoc`, `.asc` | `AddAsciiDocHandler()` | All |
| LaTeX | `OfficeIMO.Reader.Latex` | `.tex` | `AddLatexHandler()` | All |
| DocBook | `OfficeIMO.Reader.DocBook` | `.dbk`, `.docbook` | `AddDocBookHandler()` | All |
| OPML | `OfficeIMO.Reader.Opml` | `.opml` | `AddOpmlHandler()` | All |
| Jupyter notebooks | `OfficeIMO.Reader.Notebook` | `.ipynb` | `AddNotebookHandler()` | All |
| Subtitles | `OfficeIMO.Reader.Subtitles` | `.srt`, `.vtt` | `AddSubtitleHandler()` | All |
| CSV and TSV | `OfficeIMO.Reader.Csv` | `.csv`, `.tsv` | `AddCsvHandler()` | All |
| dBASE / xBase | `OfficeIMO.Reader.Dbf` | `.dbf` | `AddDbfHandler()` | All |
| JSON | `OfficeIMO.Reader.Json` | `.json` | `AddJsonHandler()` | All |
| XML | `OfficeIMO.Reader.Xml` | `.xml` | `AddXmlHandler()` | All |
| TOML | `OfficeIMO.Reader.Toml` | `.toml` | `AddTomlHandler()` | All |
| Images | `OfficeIMO.Reader.Image` | `.png`, `.jpg`, `.jpeg`, `.gif`, `.bmp`, `.tif`, `.tiff`, `.svg`, `.emf`, `.wmf`, `.ico`, `.pcx`, `.webp` | `AddImageHandler()` | All |
| ZIP archives | `OfficeIMO.Reader.Zip` | `.zip` (entries are routed to the other registered handlers) | `AddZipHandler()` | All |
| OCR enrichment | `OfficeIMO.Reader.Ocr` | Adds recognized text to candidates from other handlers | See [Optional OCR](#optional-ocr) | Not included |
| HTTP input | `OfficeIMO.Reader.Web` | Downloads with a caller-owned `HttpClient`, then reads through the registered handlers | `reader.CreateWebReader(...)` | Not included |

`OfficeIMO.Reader.All` deliberately leaves out OCR engines, process adapters, network clients and hosted providers. Detection still uses content signatures, so a known format with a wrong extension can be routed when the host enables content-first detection.

## Choose the result you need

Use `Enumerate(...)` for indexing-friendly chunks:

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Word;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddWordHandler()
    .Build();

foreach (ReaderChunk chunk in reader.Enumerate("proposal.docx")) {
    Console.WriteLine(chunk.Location.HeadingPath ?? chunk.Location.Path);
    Console.WriteLine(chunk.Markdown ?? chunk.Text);
}
```

Use `ReadDocument(...)` when the host needs the complete normalized document:

```csharp
OfficeDocumentReadResult document = reader.ReadDocument("proposal.docx");

Console.WriteLine($"{document.Pages.Count} pages or source containers");
Console.WriteLine($"{document.Tables.Count} tables");
Console.WriteLine($"{document.Assets.Count} assets");

string json = OfficeDocumentReadResultJson.Serialize(document, indented: true);
```

The rich JSON envelope currently writes schema version 7 and still reads the stable version 5 and 6 contracts. The package embeds and ships each supported JSON Schema, rejects incomplete or incompatible envelopes, and preserves structured diagnostics rather than requiring consumers to parse warning text.

## Isolated readers for services

Services and concurrent hosts should freeze their adapters, options, and processors into an isolated reader instance:

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Epub;
using OfficeIMO.Reader.Html;
using OfficeIMO.Reader.Pdf;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddPdfHandler()
    .AddHtmlHandler()
    .AddEpubHandler()
    .WithMaxConcurrentReads(4)
    .Build();

OfficeDocumentReadResult result = await reader.ReadDocumentAsync(
    "handbook.pdf",
    cancellationToken: cancellationToken);
```

`Build()` snapshots the configuration. Different reader instances can use different handlers for the same extension without leaking registration state. Async path, stream, byte, and bounded batch APIs preserve input limits, cancellation, deterministic ordering, and caller-owned stream lifetimes.

## Detection and bounds

Reader combines extension, media-type, signature, and structured container evidence. Normal reads preserve known-extension routing; hosts can opt into content-first routing for mislabeled uploads:

```csharp
ReaderDetectionResult detection = reader.Detect("upload.bin");

OfficeDocumentReadResult result = reader.ReadDocument(
    "upload.bin",
    new ReaderOptions {
        DetectionMode = ReaderDetectionMode.PreferContent,
        MaxInputBytes = 100L * 1024 * 1024,
        MaxTableRows = 1_000
    });
```

Detection reports the selected kind, extension/content evidence, confidence, media type, and mismatch state. Input bytes, folder totals, document counts, concurrency, table rows, output characters, and other expansion points have explicit bounds.

## Processing and structured extraction

Register ordered processors when every rich read should receive the same deterministic cleanup or filtering:

```csharp
OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddProcessor(new OfficeDocumentBlockNormalizationProcessor())
    .AddProcessor(new OfficeDocumentTableNormalizationProcessor())
    .AddProcessor(new OfficeDocumentLinkNormalizationProcessor())
    .WithProcessorFailureBehavior(
        OfficeDocumentProcessorFailureBehavior.ContinueWithDiagnostic)
    .Build();

OfficeDocumentStructuredExtractionResult extracted = reader.ReadDocument("contract.docx").ExtractStructured(
    options: new OfficeDocumentStructuredExtractionOptions {
        MaxRecords = 5_000,
        MaxSections = 500,
        MaxTables = 250
    });
```

The structured extractor emits bounded scalar records, heading sections, named tables, forms, chart and visual summaries, and readiness/security/OCR diagnostics. It is deterministic and does not add an AI client dependency.

## Token-aware RAG chunks

`ReadHierarchical(...)` splits large source chunks by a token budget while preserving document, page/slide/sheet, and heading ancestry:

```csharp
ReaderChunkHierarchyResult hierarchy = reader.ReadHierarchical(
    "policy.docx",
    chunkingOptions: new ReaderHierarchicalChunkingOptions {
        MaxTokens = 800,
        OverlapTokens = 80,
        MaxInputChunks = 10_000,
        MaxOutputChunks = 50_000
    });

foreach (ReaderChunk chunk in hierarchy.Chunks) {
    StoreEmbedding(chunk.Id, chunk.Text, chunk.TokenEstimate ?? 0);
}
```

The result includes deterministic leaf IDs and hashes, exact source spans, overlap/context totals, and a flattened hierarchy sidecar. Supply an `IReaderTokenCounter` when the embedding model requires its exact tokenizer.

## Optional OCR

`OfficeIMO.Ocr` owns the format-neutral `IOcrEngine` request, result, geometry, capability, and diagnostic contracts. `OfficeIMO.Reader.Ocr` executes Reader candidates with bounded count, bytes, concurrency, duration, recognized text, and geometry spans, then merges OCR output as an additional source layer without replacing native text.

- `OfficeIMO.Ocr.Process` bridges a configured executable or service through a versioned JSON protocol.
- `OfficeIMO.Ocr.Tesseract` uses a separately installed Tesseract CLI and exposes TSV line/word geometry.
- Hosts can implement `IOcrEngine` or use `DelegateOcrEngine` for another local or cloud provider.

Neither the Reader OCR integration nor either provider is a transitive dependency of `OfficeIMO.Reader.Core` or `OfficeIMO.Reader.All`. The same engine can be reused with `OfficeIMO.Pdf.Ocr` or a future format integration.

## Package and ownership boundaries

- `OfficeIMO.Reader.Core` owns the shared result, routing, reader limits, diagnostics, processing, structured extraction, and hierarchy contracts. It only carries OCR candidate metadata.
- `OfficeIMO.Ocr` owns recognition contracts; `OfficeIMO.Reader.Ocr` owns optional Reader execution and enrichment.
- Word, Excel, PowerPoint, Markdown, PDF, RTF, HTML, EPUB, Visio, and other format packages own their parsing and inspection models.
- Modular Reader packages adapt those models into the shared result.
- Storage, vector databases, AI clients, and platform-specific services belong in the consuming application or an opt-in provider.

See the [OfficeIMO.Reader.Core package README](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Reader.Core/README.md) for the complete API walkthrough and the [Reader API reference](/api/reader/) for generated type documentation.
