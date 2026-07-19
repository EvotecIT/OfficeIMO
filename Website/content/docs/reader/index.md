---
title: Reader and Extraction
description: Deterministic document extraction, rich results, structured records, RAG chunking, and optional OCR with OfficeIMO.Reader.
order: 55
---

# Reader and Extraction

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

# Broad mixed-format host
dotnet add package OfficeIMO.Reader.All
```

`OfficeIMO.Reader.Core` owns contracts, routing, plain text, processing, and custom-handler registration; it has no
format-engine dependencies. Word, Excel, PowerPoint, Markdown, Email, PDF, and the other formats are separate adapters
over their owning OfficeIMO engines.

## Choose the result you need

Use `Read(...)` for indexing-friendly chunks:

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Word;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddWordHandler()
    .Build();

foreach (ReaderChunk chunk in reader.Read("proposal.docx")) {
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

The rich JSON envelope currently writes schema version 6 and still reads the first stable version 5 contract. The package embeds and ships both JSON Schemas, rejects incomplete or incompatible envelopes, and preserves structured diagnostics rather than requiring consumers to parse warning text.

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

OfficeDocumentStructuredExtractionResult extracted = reader.ReadStructured(
    "contract.docx",
    structuredOptions: new OfficeDocumentStructuredExtractionOptions {
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

The core `IOfficeOcrEngine` contract executes OCR candidates with bounded count, bytes, concurrency, duration, recognized text, and geometry spans. OCR output is merged as an additional source layer without replacing native text.

- `OfficeIMO.Reader.Ocr.Process` bridges a configured executable or service through a versioned JSON protocol.
- `OfficeIMO.Reader.Ocr.Tesseract` uses a separately installed Tesseract CLI and exposes TSV line/word geometry.
- Hosts can implement `IOfficeOcrEngine` or use `DelegateOfficeOcrEngine` for another local or cloud provider.

Neither provider is a transitive dependency of `OfficeIMO.Reader.Core` or `OfficeIMO.Reader.All`.

## Package and ownership boundaries

- `OfficeIMO.Reader.Core` owns the shared result, routing, limits, diagnostics, processing, structured extraction, and hierarchy contracts.
- Word, Excel, PowerPoint, Markdown, PDF, RTF, HTML, EPUB, Visio, and other format packages own their parsing and inspection models.
- Modular Reader packages adapt those models into the shared result.
- Storage, vector databases, AI clients, and platform-specific services belong in the consuming application or an opt-in provider.

See the [OfficeIMO.Reader.Core package README](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Reader.Core/README.md) for the complete API walkthrough and the [Reader API reference](/api/reader/) for generated type documentation.
