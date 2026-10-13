# OfficeIMO.Reader.Core

`OfficeIMO.Reader.Core` is the dependency-light contract and orchestration package for OfficeIMO document ingestion.
It contains normalized result models, limits, deterministic routing, handler registration, processing pipelines, and
capability manifests. It does not reference Word, Excel, PowerPoint, PDF, Email, image, or other format engines.

## Install

```powershell
dotnet add package OfficeIMO.Reader.Core
```

Add only the format packages an application needs:

```powershell
dotnet add package OfficeIMO.Reader.Word
dotnet add package OfficeIMO.Reader.Email
```

Use `OfficeIMO.Reader.All` only when the complete local managed format graph is intentional.

Format handlers can register `ReaderHandlerRegistration.ReadDirectoryBundle` for directory packages with their exact registered extensions. Reader dispatches these paths as individual documents and excludes their resources from recursive folder ingestion. The handler owns bounded traversal, physical-root validation, cancellation, and snapshot hashing, and must report aggregate physical file bytes in `Source.LengthBytes`. Reader supplies the effective input-byte limit and verifies the reported size. Capability schema version 6 exposes this support through `SupportsDirectoryBundle`; ordinary directories retain the folder-ingestion path. Bundle detection reports extension/handler evidence without claiming content inspection.

## Build a reader

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Word;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddWordHandler()
    .WithMaxConcurrentReads(4)
    .Build();

OfficeDocumentReadResult document = reader.ReadDocument("Policy.docx");
```

Use `document.EnumerateContent()` to walk paragraphs and tables together in source order. Each item exposes
one `Block`, `Table`, or fallback `Chunk`, plus its effective `Location`. Tables keep their rows together and
use known source positions or an unambiguous block anchor. Content without a known position follows positioned
content in the same page, slide, or sheet. The traversal preserves the source objects; copy values when an
immutable snapshot is needed.

## Inspect review and saved content

Enable native inspection on rich document or structured reads:

```csharp
OfficeDocumentReadResult inspected = reader.ReadDocument("Review.docx", new ReaderOptions {
    Inspection = new ReaderInspectionOptions {
        MaxRecords = 4096,
        MaxCharacters = 1_000_000
    }
});

foreach (OfficeDocumentMetadataEntry entry in inspected.Metadata) {
    if (entry.Category == OfficeDocumentInspectionCategories.Comment)
        Console.WriteLine($"{entry.Value} ({entry.SourceObjectId})");
}
```

Format owners supply comments and replies, revision text, authored formulas with saved caches, hidden source state, inert package descriptors, and bounded content-safety observations where supported. Located metadata keeps native identifiers and semantic attributes through `ExtractStructured`, JSON, and nested document results. Coverage entries state the requested flags, selection, native profile, and omissions. Formula inspection reads stored values without recalculation; package and safety observations do not validate signatures or execute content.

`Inspection` is null by default. Its options are snapshotted, and record/character budgets span nested, folder, and batch reads. Exceeding an inspection or native safety admission ceiling throws terminal `ReaderResourceLimitException`; ordinary malformed inputs retain the existing recoverable per-file behavior. Unicode and concealment observations are heuristic evidence. See [migration guidance](../MIGRATION.md#typed-inspection-resource-failures) for the typed native owner exception.

The six content flags default to true when inspection is enabled: `IncludeComments`,
`IncludeRevisions`, `IncludeFormulas`, `IncludeVisibility`, `IncludePackageContent`,
and `IncludeContentSafety`. Disable individual flags to omit that work. `MaxRecords`
defaults to 4,096, including coverage records; `MaxCharacters` defaults to one million
and counts retained names, values and attributes. Both must be positive. These are
inspection-output budgets, not a bound on every allocation made by the format loader.

`IncludePackageContent` requests original-package structural observations in Word,
Excel and PowerPoint loaders, plus their inert payload/VBA descriptors. The metadata
records its physical source scope and whether that import route captured a report.
Presence of a macro or signature part does not establish execution or signature
validity. Format-specific parser limits and qualified profiles still apply.

Inspection uses the existing document metadata transport and does not introduce a
new schema version or a dependency on FileInspectorX. The format owners and Reader
adapters remain independently usable; mapping their evidence into another product
and selecting that product's package graph are separate integration decisions.

## Find content by page

Page-aware reading stays on `OfficeDocumentReadResult`; it is not a separate conversion path. A format adapter
populates `Pages`, and Reader Core provides shared location, search, and page-scoped Markdown helpers:

```csharp
OfficeDocumentSearchResult matches = document.Search("retention period");

foreach (OfficeDocumentSearchHit hit in matches.Hits) {
    Console.WriteLine(hit.Block.Text);
    foreach (OfficeDocumentPageLocation location in hit.Pages) {
        Console.WriteLine(location.Display); // for example: Page 5 of 20
    }
}

string pageMarkedMarkdown = document.ToPageMarkedMarkdown();
```

Page boundaries are not equally authoritative in every source format:

| Format | Page provenance | Reader behavior |
| --- | --- | --- |
| PDF | `Native` | Uses fixed pages and source geometry from the PDF logical model. |
| Word | `Computed` | Opt-in best-effort pagination through the OfficeIMO.Word layout engine. |
| RTF | `ExplicitBreak` | Opt-in reconstruction from explicit/saved page and section-break hints; automatic overflow is not calculated. |

Use `document.GetPageProvenance()` when page accuracy affects citations. `GetPageMarkdown()` returns separate
page values, while `ToPageMarkedMarkdown()` produces one portable Markdown string with HTML page markers.
The original document-wide `Markdown`, `Blocks`, and `Chunks` remain available on the same result.

For dependency-free plain text and an explicit unknown-payload fallback:

```csharp
OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddPlainTextHandlers()
    .Build();
```

`OfficeDocumentReader.Default` intentionally has no format handlers. This keeps Core honest: adding a format is an
explicit package and builder decision, while every built reader remains immutable and instance-scoped.

## Incremental ingestion

Use `EnumerateChunks` when a consumer can process and discard one chunk at a time:

```csharp
var reader = new OfficeDocumentReaderBuilder().AddPlainTextHandlers().Build();
var options = ReaderOptions.CreateSafeIngestion();
options.ComputeHashes = false;

foreach (ReaderChunk chunk in reader.EnumerateChunks("large.txt", options)) {
    IndexChunk(chunk);
}
```

`EnumerateChunksAsync` provides asynchronous pull and backpressure on .NET 8 and .NET 10.
Plain text supports paths and forward-only streams. CSV and Excel provide incremental path output through
their existing row-extraction engines. Excel still opens the owning workbook model and metadata; incremental
output avoids building the complete Reader result graph. Other handlers use a materialized fallback.
Check `SupportsIncrementalPath` and `SupportsIncrementalStream` in `GetCapabilities()`.

Document processors require a complete result and therefore use the materialized path. Source hashing reads
the entire source before output; forward-only inputs use a materialized fallback to preserve the source hash.
Disable `ComputeHashes` for earliest first-chunk delivery. Dispose the
enumerator when stopping early. A seekable stream reads from the beginning and restores its original position;
a forward-only stream starts at its current position. The caller keeps stream ownership and must keep its
input stable during enumeration. Content-first routing uses the snapshot-based materialized contract.

## Operation budgets and decoding

`ReaderOptions.CreateSafeIngestion()` sets a 64 MiB input limit and finite aggregate budgets for chunks,
characters, blocks, assets and nested content. Its values are a starting policy that applications can tighten:

```csharp
var options = ReaderOptions.CreateSafeIngestion();
options.ResourceLimits!.MaxChunks = 2_000;
options.ResourceLimits.MaxNestedInputBytes = 32L * 1024 * 1024;
options.ResourceLimits.MaxNestedDepth = 3;
options.ThrowOnInvalidTextBytes = true;
```

`ResourceLimits` is optional; unset budgets remain unlimited. Limits span the root and its nested reads,
all files in one folder operation, and all workers in one batch. Exceeding them throws
`ReaderResourceLimitException` and ends the operation, including resilient folder and batch routes.
`MaxChunkCharacters` counts both `Text` and `Markdown`. Nested input budgets count decoded entries,
including intermediate archive payloads. Shared chunk/block objects and asset byte arrays count once.
Rich-model limits are checked when the engine returns its model; retain the owning format's parser,
archive, decompression and asset limits to control allocations inside that engine.

`MaxChars` is a best-effort projection target. Atomic Markdown blocks and tables can exceed it with a warning;
Reader preserves their complete content. Use `ResourceLimits.MaxChunkCharacters` for a terminal operation limit.
Folder and detailed path reads run the same configured processors as single-document reads.

Plain-text handlers recognize UTF-8, UTF-16 and UTF-32 BOMs. Without a BOM they use UTF-8 or the supplied
`TextEncoding`. Replacement decoding emits a warning; `ThrowOnInvalidTextBytes` rejects malformed bytes.
BOMs override `TextEncoding`. These options apply to plain-text handlers; other formats use their own
codec policies. Chunk boundaries preserve Unicode surrogate pairs, and line breaks normalize to LF.

## Nested results and qualified capabilities

ZIP and email reads keep each delegated source's rich result in `NestedDocuments`, alongside the existing
flattened chunks. Each item supplies its virtual `Path` and complete `Document`, including available links,
forms, metadata, diagnostics and assets. Identifiers inside a child result remain local to that child.
Asset payload bytes stay in memory when requested and remain excluded from JSON transport.

Document transport schema version 14 adds Publisher input identity (`ReaderInputKind.Publisher`, 30).
Version 13 adds DBF input identity (`ReaderInputKind.Dbf`, 29).
Version 12 adds DjVu identity (`ReaderInputKind.DjVu`, 28); version 11 adds CHM identity
(`ReaderInputKind.Chm`, 27). Version 10 adds XPS/OpenXPS identity and explicit native
`ReaderLocation.LogicalOrder` across physical containers; version 9 includes recursive
`nestedDocuments`. Versions 5 through 13 remain readable. Load the matching artifact
with `OfficeDocumentReadResultSchema.GetJsonSchema(version)`. Versions below 9 cannot
carry nested documents, below 10 cannot carry XPS input kinds, below 11 cannot carry
CHM input kinds, below 12 cannot carry DjVu input kinds, below 13 cannot carry DBF input kinds,
and below 14 cannot carry Publisher input kinds.

Capability manifest version 6 includes incremental route flags and per-extension `FormatQualifications`.
A qualification records a format ID, extraction maturity, profile, preservation, limitations and evidence
references. Word, Excel and modern PowerPoint reuse their owner's format IDs. Other undeclared profiles
remain `Unqualified`; an available handler is not a claim of complete format preservation. Applications can
supply immutable `ReaderFormatQualification` values when registering a custom handler.

For model-specific token budgets, pass a `ReaderDelegateTokenCounter` to `ReaderHierarchicalChunkingOptions`:

```csharp
var chunking = new ReaderHierarchicalChunkingOptions {
    MaxTokens = 800,
    TokenCounter = new ReaderDelegateTokenCounter("my-model:vocabulary-version", tokenizer.CountTokens)
};
```

The optional range callback counts the exact prefix-plus-source range without a temporary substring.
The application owns tokenizer dependencies and thread safety. The built-in counter estimates one token
per four UTF-16 characters; it does not represent a model vocabulary or guarantee a model token budget.

## Package selection

| Need | Package |
| --- | --- |
| Contracts, routing, processors, schemas | `OfficeIMO.Reader.Core` |
| Word only | `OfficeIMO.Reader.Word` |
| Excel only | `OfficeIMO.Reader.Excel` |
| PowerPoint only | `OfficeIMO.Reader.PowerPoint` |
| Markdown only | `OfficeIMO.Reader.Markdown` |
| Email artifacts, stores, and OAB | `OfficeIMO.Reader.Email` |
| PDF only | `OfficeIMO.Reader.Pdf` |
| Every local managed handler | `OfficeIMO.Reader.All` |

Other `OfficeIMO.Reader.*` packages follow the same rule: Core plus the format's owning engine. OCR processes,
network clients, hosted providers, and native tools remain explicit host choices and are not composed by All.

## Stable contracts

File reads retain the same source identity and timestamps through synchronous, asynchronous and batch
routes. Container result envelopes describe the outer input; member chunks retain their own provenance
through document processors. Folder byte limits charge the physical files accepted for parsing.
Native async file handlers are awaited directly. Synchronous file handlers run on a worker through the
reader's concurrency gate; their file-specific behavior also applies to async and batch reads.

Table exports accept cancellation while scanning and writing rows. Pass a token to
`reader.ExportTables(tables, cancellationToken: token)` for CSV, Markdown and JSON together, or use
`table.ToCsv(token)`, `table.ToMarkdownTable(token)` and `table.ToJson(indented: true, cancellationToken: token)` separately.

- `ReaderOptions` and format-neutral input/processing limits
- `ReaderChunk` and the schema-versioned `OfficeDocumentReadResult`
- tables, pages, visuals, assets, links, forms, metadata, and diagnostics
- sync/async path, stream, byte-array, folder, and batch ingestion
- deterministic capability manifests with `OfficeIMO` versus `Custom` handler origins
- bounded nested-content delegation between configured handlers

Public namespaces remain `OfficeIMO.Reader`; the `.Core` name describes package and assembly ownership.
