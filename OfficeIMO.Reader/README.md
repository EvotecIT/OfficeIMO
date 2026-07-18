# OfficeIMO.Reader - document extraction facade

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Reader)](https://www.nuget.org/packages/OfficeIMO.Reader)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Reader?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Reader)

`OfficeIMO.Reader` is a read-only facade for deterministic document extraction. It normalizes supported source files into `ReaderChunk` objects for search, indexing, chat, RAG, migration, and review workflows.

Built-in email-family handling includes EML, MSG, OFT, TNEF, Mbox, standalone iCalendar (`.ics`/`.vcs`), and
standalone vCard (`.vcf`/`.vcard`). Calendar and contact extraction delegates to the public lossless engines in
`OfficeIMO.Email`; optional PST, OST, OLM, and EMLX stores remain in `OfficeIMO.Reader.EmailStore`.

## Install

```powershell
dotnet add package OfficeIMO.Reader
```

## Quick start

```csharp
using OfficeIMO.Reader;

foreach (var chunk in OfficeDocumentReader.Default.Read(@"C:\Docs\Policy.docx")) {
    Console.WriteLine(chunk.Id);
    Console.WriteLine(chunk.Location.HeadingPath);
    Console.WriteLine(chunk.Markdown ?? chunk.Text);
}
```

When the caller only needs the portable Markdown projection, use the thin conversion helper:

```csharp
string markdown = await OfficeDocumentReader.Default.ConvertToMarkdownAsync(
    @"C:\Docs\Policy.docx",
    cancellationToken: cancellationToken);
```

`ConvertToMarkdown(...)` and `ConvertToMarkdownAsync(...)` support file paths, streams, and byte arrays. They return the `Markdown` emitted by the same rich `ReadDocument(...)` pipeline, or an empty string when a handler emits no Markdown. Use the rich read APIs when metadata, blocks, tables, assets, or diagnostics are needed.

## Streams and folders

```csharp
using OfficeIMO.Reader;

using var stream = File.OpenRead(@"C:\Docs\Policy.docx");
var chunksFromStream = OfficeDocumentReader.Default.Read(stream, "Policy.docx").ToList();

var folderChunks = OfficeDocumentReader.Default.ReadFolder(
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

## What it reads

Built-in and modular adapters can extract:

- Word (`.docx`, `.docm`, `.doc`) as Markdown chunks.
- Excel (`.xlsx`, `.xlsm`, `.xls`) as table chunks and optional Markdown previews.
- PowerPoint (`.pptx`, `.pptm`, `.ppt`, `.pot`, `.pps`) as slide-aligned chunks, optionally including notes. Binary files use the OfficeIMO.PowerPoint projection model and expose import diagnostics as warnings.
- Markdown (`.md`, `.markdown`) as parser-aware heading chunks.
- EML/MIME, Outlook MSG/MAPI, TNEF/`winmail.dat`, and mbox as message, body, attachment, and embedded-item chunks. Rich results include attachment assets and typed Outlook metadata.
- OpenDocument (`.odt`, `.ods`, `.odp`), PDF, RTF, Visio, HTML, CSV/TSV, JSON, XML, YAML, EPUB, ZIP, standalone images, Jupyter notebooks, SRT/WebVTT subtitles, and structured text through modular adapter packages.

Use `OpenPassword` for password-protected Open XML or legacy binary Office input. Binary PowerPoint images
are returned through the same rich-result asset contract as PPTX images:

```csharp
OfficeDocumentReadResult deck = DocumentReader.ReadDocument(
    "protected.ppt",
    new ReaderOptions {
        OpenPassword = "open-password",
        IncludePowerPointNotes = true
    });

Console.WriteLine($"{deck.Pages.Count} slides, {deck.Assets.Count} assets");
foreach (OfficeDocumentDiagnostic diagnostic in deck.Diagnostics) {
    Console.WriteLine($"{diagnostic.Code}: {diagnostic.Message}");
}
```

## Modular adapters

For services and concurrent hosts, build an isolated reader with only the adapters you need:

```csharp
using OfficeIMO.Reader.Csv;
using OfficeIMO.Reader.Epub;
using OfficeIMO.Reader.Html;
using OfficeIMO.Reader.Image;
using OfficeIMO.Reader.Json;
using OfficeIMO.Reader.Notebook;
using OfficeIMO.Reader.OpenDocument;
using OfficeIMO.Reader.Pdf;
using OfficeIMO.Reader.Rtf;
using OfficeIMO.Reader.Subtitles;
using OfficeIMO.Reader.Visio;
using OfficeIMO.Reader.Xml;
using OfficeIMO.Reader.Yaml;
using OfficeIMO.Reader.Zip;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddCsvHandler()
    .AddEpubHandler()
    .AddHtmlHandler()
    .AddImageHandler()
    .AddJsonHandler()
    .AddNotebookHandler()
    .AddOpenDocumentHandler()
    .AddPdfHandler()
    .AddRtfHandler()
    .AddSubtitleHandler()
    .AddVisioHandler()
    .AddXmlHandler()
    .AddYamlHandler()
    .AddZipHandler()
    .Build();

var chunks = reader.Read(@"C:\Docs\data.json").ToList();
```

`Build()` freezes the handler configuration. The resulting `OfficeDocumentReader` is safe to reuse across concurrent reads and is unaffected by later builder changes. `OfficeDocumentReader.Default` is the simple built-in-only instance; modular formats are configured explicitly on `OfficeDocumentReaderBuilder`, so one reader cannot change another reader or process-wide state.

When the application wants every local adapter, `OfficeIMO.Reader.All` provides the same explicit, instance-scoped composition in one call:

```powershell
dotnet add package OfficeIMO.Reader.All
```

```csharp
using OfficeIMO.Reader.All;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddAllOfficeIMOHandlers()
    .Build();
```

The preset contains no parser or provider logic. It excludes OCR engines, process adapters, network clients, and hosted providers; those remain explicit host choices. Standalone images expose header metadata, source assets, and optional OCR candidates without running OCR. Notebook cells are never executed, and subtitle support reads local SRT/WebVTT text only.

### Explicit bounded web transport

Install `OfficeIMO.Reader.Web` only in hosts that intentionally allow network retrieval:

```powershell
dotnet add package OfficeIMO.Reader.Web
```

```csharp
using OfficeIMO.Reader.Web;

OfficeDocumentWebReader webReader = reader.CreateWebReader(
    httpClient,
    new ReaderWebOptions {
        AllowedHosts = new[] { "docs.example.com" },
        MaxResponseBytes = 16L * 1024 * 1024
    });

OfficeDocumentReadResult remote = await webReader.ReadDocumentAsync(
    new Uri("https://docs.example.com/guide.html"),
    cancellationToken: cancellationToken);
```

The web package performs bounded HTTP(S) GET retrieval and routes the bytes through the same configured Reader handlers and processors. It captures its options, bounds bytes, time, and concurrent operations, blocks obvious private literal targets by default, and omits query strings from metadata by default. The caller owns the `HttpClient` and its authentication, DNS, redirect, proxy, certificate, retry, and connection policy. It is deliberately absent from `OfficeIMO.Reader.All` and the global tool, so neither surface gains implicit network access.

For scripts and shell pipelines, install the dependency-bounded global tool:

```powershell
dotnet tool install --global OfficeIMO.Reader.Tool
officeimo-reader read policy.docx --format markdown --output policy.md
officeimo-reader capabilities --format json
```

`officeimo-reader folder` provides deterministic, bounded folder conversion with configurable concurrency and optional asset directories. The tool emits Markdown or the stable Reader v5 JSON envelope and does not configure OCR or hosted providers.

## Async and bounded batches

Configure the instance-wide async limit once, then use the same reader for individual or batched work:

```csharp
using OfficeIMO.Reader;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .WithMaxConcurrentReads(4)
    .Build();

OfficeDocumentReadResult document = await reader.ReadDocumentAsync(
    @"C:\Docs\Policy.docx",
    cancellationToken: cancellationToken);

IReadOnlyList<OfficeDocumentReadResult> documents = await reader.ReadDocumentsAsync(
    Directory.EnumerateFiles(@"C:\Docs", "*.docx"),
    batchOptions: new ReaderBatchOptions {
        MaxDegreeOfParallelism = 3,
        MaxDocuments = 500
    },
    cancellationToken: cancellationToken);
```

`ReadDocumentsAsync(...)` starts no more than the configured degree of parallelism, rejects input beyond `MaxDocuments`, cancels sibling workers after a failure, and returns results in the same order as the input paths. Handlers can provide `ReadDocumentPathAsync` or `ReadDocumentStreamAsync` for native asynchronous work. Existing synchronous format engines remain compatible and are scheduled through the instance's bounded worker gate.

## Ordered document processors

Processors are opt-in transformations over `OfficeDocumentReadResult`. The builder freezes their order with the handler configuration, and the reader applies them to `Read(...)`, `ReadAsync(...)`, rich document reads, JSON reads, structured reads, and `ReadDocumentsAsync(...)`:

```csharp
OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddProcessor(new OfficeDocumentBlockNormalizationProcessor())
    .AddProcessor(new OfficeDocumentTableNormalizationProcessor())
    .AddProcessor(new OfficeDocumentLinkNormalizationProcessor())
    .AddProcessor(new OfficeDocumentArtifactClassificationProcessor())
    .AddProcessor(new OfficeDocumentAssetFilterProcessor(
        asset => asset.LengthBytes <= 5 * 1024 * 1024))
    .WithProcessorFailureBehavior(
        OfficeDocumentProcessorFailureBehavior.ContinueWithDiagnostic)
    .Build();

OfficeDocumentReadResult document = reader.ReadDocument(@"C:\Docs\Policy.docx");
```

The built-ins normalize shared blocks, list and heading levels, tables, and links; classify repeated page-boundary text; and filter assets together with dependent OCR candidates. They do not run unless registered, so an unconfigured reader preserves existing output. Add host behavior with `OfficeDocumentProcessorBase`, `IOfficeDocumentProcessor`, or `DelegateOfficeDocumentProcessor`.

The default failure policy is `Throw`. `ContinueWithDiagnostic` records `processor-failed` and runs later steps, while `StopWithDiagnostic` records the failure and marks later steps as skipped. Call `ProcessDocument(...)` or `ProcessDocumentAsync(...)` directly when the host needs the per-step `OfficeDocumentProcessingResult`. Processor instances attached to a shared reader must be safe for concurrent calls and should avoid retaining per-document state.

Folder enumeration remains a chunk/source traversal surface and does not run rich-document processors automatically. Use `ReadDocumentsAsync(...)` when every file must pass through the configured document pipeline.

## Bounded structured extraction

Structured extraction projects the shared document model into deterministic scalar records, heading sections, named tables, typed forms, and diagnostics without adding an AI client or format-specific dependency:

```csharp
OfficeDocumentStructuredExtractionResult extracted = reader.ReadStructured(
    @"C:\Docs\Policy.docx",
    structuredOptions: new OfficeDocumentStructuredExtractionOptions {
        MaxRecords = 5_000,
        MaxSections = 500,
        MaxSectionCharacters = 50_000,
        MaxTables = 250,
        MaxForms = 1_000,
        MaxDiagnostics = 500
    });

foreach (OfficeDocumentStructuredRecord record in extracted.Records) {
    Console.WriteLine($"{record.Category}: {record.Name} = {record.Value}");
}

string json = extracted.ToJson(indented: true);
```

The extractor recognizes metadata, forms, two-column key/value tables, Path/Type/Value tables, Visio Shape Data, chart summaries, table and visual quality, chunk readiness, and security/OCR/limit diagnostics. Each collection and section body has an explicit positive limit; reaching a bound adds a structured limit diagnostic instead of silently growing output. The non-generic schema-friendly shape is versioned independently from the stable `OfficeDocumentReadResult` transport. It is currently a serialization and host-integration contract; a generic `ExtractStructured<T>()` mapper and model-assisted extraction are intentionally outside the core package.

## Token-aware hierarchical chunks

`ReadHierarchical(...)` keeps `ReaderChunk` as the embedding leaf while adding document, page/slide/sheet, and heading nodes around it:

```csharp
ReaderChunkHierarchyResult hierarchy = reader.ReadHierarchical(
    @"C:\Docs\Policy.docx",
    chunkingOptions: new ReaderHierarchicalChunkingOptions {
        MaxTokens = 800,
        OverlapTokens = 80,
        MaxInputChunks = 10_000,
        MaxOutputChunks = 50_000,
        MaxHierarchyDepth = 32,
        MaxContextCharacters = 4_096,
        IncludeContextInText = true
    });

foreach (ReaderChunk chunk in hierarchy.Chunks) {
    StoreEmbedding(chunk.Id, chunk.Text, chunk.TokenEstimate ?? 0);
}

string sidecarJson = hierarchy.ToJson(indented: true);
```

The result records the selected token counter, original/output/overlap/context token totals, exact source character spans, deterministic leaf IDs and hashes, and a flattened hierarchy with explicit parent/child IDs. Splits prefer paragraph, line, sentence, and whitespace boundaries but enforce the token maximum even when a source block is larger. Structured tables, visuals, forms, actions, diagnostics, and warnings stay on the first segment so overlap does not duplicate sidecars.

The dependency-free default uses the existing four-characters-per-token heuristic. Supply a deterministic, thread-safe `IReaderTokenCounter` when the embedding model needs its exact tokenizer. A hierarchy represents one source document and rejects mixed-source chunk collections. Context and input/output/depth bounds emit structured diagnostics; invalid token budgets or counters fail explicitly.

This is an opt-in projection with its own schema id/version. It does not add fields to `ReaderChunk` or change normal reads. Static and instance sync/async file, stream, and byte overloads are available; instance reads apply configured processors before chunking.

## Stable document transport

`OfficeDocumentReadResult` schema version 5 is the first stable JSON transport contract. Version 6 adds the `Calendar` and `VCard` kinds without changing the closed version 5 enum; versions 1 through 4 remain deliberately unsupported.

```csharp
OfficeDocumentReadResult document = reader.ReadDocument(@"C:\Docs\Policy.docx");
string json = OfficeDocumentReadResultJson.Serialize(document, indented: true);

OfficeDocumentReadResult restored = OfficeDocumentReadResultJson.Deserialize(json);
Console.WriteLine(restored.SchemaVersion); // 6

string jsonSchema = OfficeDocumentReadResultSchema.GetJsonSchema();
```

The package embeds and ships both `schemas/officeimo.document.read-result.v5.schema.json` and `schemas/officeimo.document.read-result.v6.schema.json`. Deserialization accepts stable versions 5 and 6, normalizes accepted payloads to the current model, rejects unknown top-level members and incomplete envelopes, and returns a typed `OfficeDocumentReadResultSchemaException` for incompatible versions. `GetJsonSchema()` returns the current schema; `GetJsonSchema(version)` returns a supported historical artifact. Published core and modular Reader packages run NuGet package validation against their current public versions, so accidental API breaks fail packaging.

## Content detection and structured diagnostics

`Detect(...)` and `DetectAsync(...)` report extension and content evidence without reducing the answer to one opaque enum:

```csharp
ReaderDetectionResult detection = reader.Detect(@"C:\Inbox\upload.bin");

Console.WriteLine($"Selected: {detection.Kind}");
Console.WriteLine($"Extension: {detection.ExtensionKind}");
Console.WriteLine($"Content: {detection.ContentKind} ({detection.ContentConfidence})");
Console.WriteLine(string.Join(", ", detection.Evidence));
```

Standalone detection prefers medium- or high-confidence content evidence. Normal reads use `ContentWhenUnknown` by default, which preserves known-extension behavior while identifying unknown uploads. Opt into mislabeled-file routing when needed:

```csharp
OfficeDocumentReadResult result = reader.ReadDocument(
    @"C:\Inbox\actually-markdown.txt",
    new ReaderOptions { DetectionMode = ReaderDetectionMode.PreferContent });
```

Seekable stream probes restore the original position. Prefix probing is capped at 4 MiB, and ZIP-based Office/Visio/EPUB inspection walks at most 4,096 local entries without decompressing archive payloads. Detection mismatches, detected unknown inputs, parser failures, limits, truncation, unsupported content, read failures, and OCR readiness are exposed through `OfficeDocumentDiagnostic` with stable `Category`, `Code`, `Source`, `IsRecoverable`, and `Attributes` fields.

## Rich built-in document mappings

Word, Excel, and PowerPoint rich reads use the format packages' public inspection models instead of reparsing Open XML in the Reader facade:

```csharp
OfficeDocumentReadResult word = OfficeDocumentReader.Default.ReadDocument("policy.docx");
OfficeDocumentReadResult workbook = OfficeDocumentReader.Default.ReadDocument("forecast.xlsx");
OfficeDocumentReadResult deck = OfficeDocumentReader.Default.ReadDocument("briefing.pptx");

Console.WriteLine($"{word.Blocks.Count} semantic Word blocks");
Console.WriteLine($"{workbook.Tables.Count} named Excel tables");
Console.WriteLine($"{deck.Visuals.Count} PowerPoint chart snapshots");
```

The Word mapping preserves headings, lists, links, tables, and header/footer content. Excel exposes worksheet locations, formal tables, cell links, formula/comment/named-range counts, and workbook properties. PowerPoint exposes slide-local blocks and tables, run and shape links, chart snapshots, and point-based geometry. Modular HTML, EPUB, RTF, and Visio adapters register the same native v5 result surface when their packages are installed.

## Optional OCR execution

The core defines `IOfficeOcrEngine` and runs caller-supplied engines without taking an OCR or cloud SDK dependency. Execution resolves candidate assets, validates payload hashes and media types, and bounds candidate count, per-candidate and total bytes, concurrency, duration, recognized characters, and detailed spans:

```csharp
var engine = new DelegateOfficeOcrEngine(
    "host-vision-service",
    async (request, cancellationToken) => {
        HostVisionResponse response = await visionClient.RecognizeAsync(
            request.Payload,
            request.Language,
            cancellationToken);

        return new OfficeOcrEngineResult {
            Text = response.Text,
            Confidence = response.Confidence,
            Language = response.Language,
            Provider = "host-vision-service"
        };
    });

OfficeDocumentReadResult source = OfficeDocumentReader.Default.ReadDocument("scanned.pdf");
OfficeDocumentOcrExecutionResult execution = await source.ApplyOcrAsync(
    engine,
    new OfficeDocumentOcrExecutionOptions {
        MaxCandidates = 25,
        MaxDegreeOfParallelism = 2
    });
```

To run the same frozen configuration automatically for every rich read, register `new OfficeDocumentOcrProcessor(engine, executionOptions)` with `OfficeDocumentReaderBuilder.AddProcessor(...)`. Engines that do not advertise concurrent-request support are serialized per engine instance even when several reader operations run concurrently.

`execution.Document` contains merged `ocr-text` blocks/chunks while unresolved candidates and diagnostics remain intact. `execution.Recognitions` carries optional line, word, and character spans with confidence, language, bounding boxes, and coordinate units. Detailed provider spans intentionally stay outside the stable versioned transport; the merged text and trace metadata use the existing schema.

Use `OfficeIMO.Reader.Ocr.Process` for a versioned local executable/service bridge or `OfficeIMO.Reader.Ocr.Tesseract` for an installed Tesseract CLI. Neither package is pulled transitively by `OfficeIMO.Reader`.

## Host examples

### Capability discovery

```csharp
using OfficeIMO.Reader;

var reader = new OfficeDocumentReaderBuilder().Build();
var capabilities = reader.GetCapabilities();
foreach (var capability in capabilities) {
    Console.WriteLine($"{capability.Id}: {string.Join(", ", capability.Extensions)}");
}

string manifestJson = reader.GetCapabilityManifestJson();
```

### Register a custom handler

```csharp
using OfficeIMO.Reader;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddHandler(new ReaderHandlerRegistration {
        Id = "custom-audit",
        DisplayName = "Custom audit reader",
        Kind = ReaderInputKind.Text,
        Extensions = new[] { ".auditx" },
        ReadPath = (path, options, cancellationToken) => {
            string text = File.ReadAllText(path);
            return new[] {
                new ReaderChunk {
                    Id = "audit:1",
                    Kind = ReaderInputKind.Text,
                    Text = text,
                    Location = new ReaderLocation { Path = path }
                }
            };
        }
    })
    .Build();
```

Handlers that already expose a structured document model can register rich result delegates instead of rebuilding that model as chunks first:

```csharp
OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddHandler(new ReaderHandlerRegistration {
        Id = "custom-rich-reader",
        DisplayName = "Custom rich reader",
        Kind = ReaderInputKind.Text,
        Extensions = new[] { ".rich" },
        ReadDocumentPath = (path, options, cancellationToken) => ReadRichDocument(path),
        ReadDocumentStream = (stream, sourceName, options, cancellationToken) => ReadRichDocument(stream, sourceName)
    })
    .Build();
```

`reader.ReadDocument(...)` dispatches directly to these delegates. Existing `reader.Read(...)` calls remain usable by projecting the returned result's `Chunks` collection. A handler may continue to register `ReadPath` and `ReadStream` when chunk production is its native contract.

Rich handlers return `OfficeDocumentReadResult` directly and can populate the common v5 source, chunks, assets, diagnostics, and baseline collections before adding format-owned blocks, pages, tables, links, forms, visuals, and metadata.

## Host contracts

- `ReaderOptions` controls chunk size, table row limits, footnotes/notes, passwords, Excel ranges, Markdown heading chunking, hashes, and input budgets.
- `ReaderFolderOptions` controls recursion, file limits, byte limits, reparse-point handling, and deterministic folder order.
- `OfficeDocumentReader.GetCapabilities()` and `GetCapabilityManifestJson()` expose the frozen configuration of that reader instance.
- Capability records distinguish basic path/stream support from native rich-result support through `SupportsDocumentPath` and `SupportsDocumentStream`.
- `SupportsAsyncPath` and `SupportsAsyncStream` identify handlers with native asynchronous delegates; false means the async facade uses the bounded synchronous fallback.
- `OfficeDocumentReader.Detect(...)` / `DetectAsync(...)` expose bounded extension/content evidence, confidence, media type, and mismatch state.
- `OfficeDocumentDiagnostic` carries stable categories, codes, sources, recoverability, and attributes so hosts do not need to parse warning text.
- `OfficeDocumentReadResultJson` reads stable schema versions 5 and 6 and writes version 6 by default; `OfficeDocumentReadResultSchema.GetJsonSchema()` exposes the packaged current schema artifact.
- `OfficeDocumentProcessorPipeline` freezes ordered processors and reports each completed, failed, or skipped step.
- `OfficeDocumentStructuredExtractor` produces bounded non-generic records, sections, named tables, forms, and diagnostics without AI dependencies.
- `ReaderHierarchicalChunker` produces token-bounded `ReaderChunk` leaves, exact overlap spans, and document/container/heading hierarchy nodes.
- `OfficeDocumentReaderBuilder.AddHandler(...)` is the recommended custom-handler path for services and concurrent hosts.
- Reader configuration and custom handler registration stay instance-scoped and are frozen by `OfficeDocumentReaderBuilder.Build()`.

## Boundaries

- `OfficeIMO.Reader` owns the shared extraction contract and built-in facade.
- Source-specific parsing belongs in the source package or modular adapter.
- Email RTF transport decoding belongs to `OfficeIMO.Email`; semantic extraction of an RTF-only email body is delegated to the registered `OfficeIMO.Reader.Rtf` adapter.
- Adapters should use `ReaderInputLimits` so input size and stream behavior stays consistent.
- AI or database storage belongs in the consuming application.

## Performance evidence

`OfficeIMO.Reader.Benchmarks` covers rich extraction across all built-in and modular formats, bounded content detection, schema serialization/deserialization, and parser-versus-Reader isolation. Run it from the repository root with:

```powershell
dotnet run --project OfficeIMO.Reader.Benchmarks/OfficeIMO.Reader.Benchmarks.csproj -c Release -f net8.0
```

The benchmark corpus is generated deterministically before measurement. Benchmark artifacts are machine-specific and remain untracked; durable baseline notes record the runtime, hardware, corpus, and relevant comparisons.

## Related packages

- [OfficeIMO.Email](../OfficeIMO.Email/README.md)
- [OfficeIMO.Reader.Pdf](../OfficeIMO.Reader.Pdf/README.md)
- [OfficeIMO.Reader.Rtf](../OfficeIMO.Reader.Rtf/README.md)
- [OfficeIMO.Reader.Visio](../OfficeIMO.Reader.Visio/README.md)
- [OfficeIMO.Reader.Html](../OfficeIMO.Reader.Html/README.md)
- [OfficeIMO.Reader.Csv](../OfficeIMO.Reader.Csv/README.md)
- [OfficeIMO.Reader.Ocr.Process](../OfficeIMO.Reader.Ocr.Process/README.md)
- [OfficeIMO.Reader.Ocr.Tesseract](../OfficeIMO.Reader.Ocr.Tesseract/README.md)
- [OfficeIMO.Reader.Json](../OfficeIMO.Reader.Json/README.md)
- [OfficeIMO.Reader.OpenDocument](../OfficeIMO.Reader.OpenDocument/README.md)
- [OfficeIMO.Reader.Xml](../OfficeIMO.Reader.Xml/README.md)
- [OfficeIMO.Reader.Yaml](../OfficeIMO.Reader.Yaml/README.md)
- [OfficeIMO.Reader.Epub](../OfficeIMO.Reader.Epub/README.md)
- [OfficeIMO.Reader.Zip](../OfficeIMO.Reader.Zip/README.md)

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`.
- License: MIT.
- Repository: [EvotecIT/OfficeIMO](https://github.com/EvotecIT/OfficeIMO)

## Dependency footprint

- **External:** `System.Text.Json` for schema/result serialization.
- **OfficeIMO:** Native Word, Excel, PowerPoint, email/iCalendar/vCard, Markdown, PDF, and Drawing engines are reused directly; optional formats stay in adapter packages.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.
