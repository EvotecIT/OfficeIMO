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

Format handlers can register `ReaderHandlerRegistration.ReadDirectoryBundle` for directory packages with their exact registered extensions. Reader dispatches these paths as individual documents and excludes their resources from recursive folder ingestion. The handler owns bounded traversal, physical-root validation, cancellation, and snapshot hashing, and must report aggregate physical file bytes in `Source.LengthBytes`. Reader supplies the effective input-byte limit and verifies the reported size. Capability schema version 7 exposes this support through `SupportsDirectoryBundle`; ordinary directories retain the folder-ingestion path. Bundle detection reports extension/handler evidence without claiming content inspection.

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
using OfficeIMO.Reader;
using OfficeIMO.Reader.Word;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddWordHandler().Build();

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
using OfficeIMO.Reader;
using OfficeIMO.Reader.Word;

OfficeDocumentReadResult document = new OfficeDocumentReaderBuilder().AddWordHandler().Build().ReadDocument("Policy.docx");

OfficeDocumentSearchResult matches = document.Search("retention period");

foreach (OfficeDocumentSearchHit hit in matches.Hits) {
    Console.WriteLine(hit.Block.Text);
    foreach (OfficeDocumentPageLocation location in hit.Pages) {
        Console.WriteLine(location.Display); // for example: Page 5 of 20
    }
}

string pageMarkedMarkdown = document.ToPageMarkedMarkdown();
string text = document.GetText();
IReadOnlyList<OfficeDocumentPageText> pageText = document.GetPageText();
```

`Search` matches literal text by default. Ranked mode returns one hit per block, best first, scored with BM25 over
Unicode (UAX #29) words with NFKC normalization and case folding; words in the heading path weigh double. Queries
accept `"quoted phrases"`, `prefix*`, `OR` and parentheses, and adjacent words must all match. A query with no
words, or no ranked match, falls back to literal search, which `Ranked` on the result reports.

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Word;

OfficeDocumentReadResult document = new OfficeDocumentReaderBuilder().AddWordHandler().Build().ReadDocument("Policy.docx");

OfficeDocumentSearchResult ranked = document.Search("\"retention period\" OR archiv*",
    new OfficeDocumentSearchOptions { Mode = OfficeDocumentSearchMode.Ranked });
foreach (OfficeDocumentSearchHit hit in ranked.Hits) {
    Console.WriteLine($"{hit.Score:F3} {hit.Block.Text}"); // hit.Highlights holds every matched range
}

// Index chunks once and query them repeatedly.
var index = new ReaderChunkSearchIndex(document.Chunks);
ReaderChunkSearchResult chunks = index.Search("retention");
```

Page boundaries are not equally authoritative in every source format:

| Format | Page provenance | Reader behavior |
| --- | --- | --- |
| PDF | `Native` | Uses fixed pages and source geometry from the PDF logical model. |
| Word | `Computed` | Opt-in best-effort pagination through the OfficeIMO.Word layout engine. |
| RTF | `ExplicitBreak` | Opt-in reconstruction from explicit/saved page and section-break hints; automatic overflow is not calculated. |

Use `document.GetPageProvenance()` when page accuracy affects citations. `GetPageMarkdown()` returns separate
page values, while `ToPageMarkedMarkdown()` produces one portable Markdown string with HTML page markers.
`GetText()` follows the same canonical source order, emits structured table cells separated by tabs,
and avoids duplicate block, chunk and table projections. `GetPageText()` returns only observed page-like
containers, with their original page metadata and provenance. These projections do not reopen the source.
The original document-wide `Markdown`, `Blocks`, and `Chunks` remain available on the same result.

For dependency-free plain text and an explicit unknown-payload fallback:

```csharp
using OfficeIMO.Reader;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddPlainTextHandlers()
    .Build();
```

For source code, `AddSourceCodeHandler()` reads common programming, scripting and configuration files
(`.cs`, `.ps1`, `.psm1`, `.py`, `.js`, `.ts`, `.java`, `.go`, `.rs`, `.c`, `.cpp`, `.h`, `.sql`, `.sh`, `.ini` and more):

```csharp
using OfficeIMO.Reader;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddPlainTextHandlers()
    .AddSourceCodeHandler()
    .Build();

OfficeDocumentReadResult script = reader.ReadDocument("deploy.ps1");
string? language = script.Metadata.FirstOrDefault(entry => entry.Id == "source-code-language")?.Value; // "powershell"
```

`Text` keeps the source verbatim with LF line breaks. `Markdown` wraps each chunk in a fenced block tagged with the
language, using a longer fence when the code itself contains backticks. Chunks split on line boundaries and carry
their start and end lines. The `source-code.document` metadata reports the language (with its display name and
whether the extension, a conventional file name or a `#!` line identified it) and the line count. Extensionless files
named by convention, such as `Dockerfile`, `Containerfile`, `Makefile`, `Jenkinsfile`, `Gemfile`, `Rakefile` and
`Vagrantfile`, are read as their language, including in folder reads. Other extensionless scripts whose first line is a
shebang such as `#!/usr/bin/env python3` are routed to this handler by content detection. Decoding follows the
plain-text rules below; input containing NUL bytes outside UTF-16 or UTF-32 is refused with
`InvalidDataException`. JSON, YAML, TOML, XML, Markdown, CSV, HTML, LaTeX and the other structured formats stay with
their own readers; this handler does not claim their extensions. `AddAllOfficeIMOHandlers()` includes it.

`OfficeDocumentReader.Default` intentionally has no format handlers. This keeps Core honest: adding a format is an
explicit package and builder decision, while every built reader remains immutable and instance-scoped.

## Incremental ingestion

Use `Enumerate` when a consumer can process and discard one chunk at a time:

```csharp
using OfficeIMO.Reader;

var reader = new OfficeDocumentReaderBuilder().AddPlainTextHandlers().Build();
var options = ReaderOptions.CreateSafeIngestion();
options.ComputeHashes = false;

foreach (ReaderChunk chunk in reader.Enumerate("large.txt", options)) {
    IndexChunk(chunk);
}

static void IndexChunk(ReaderChunk chunk) => Console.WriteLine(chunk.Id);
```

`EnumerateAsync` provides asynchronous pull and backpressure on .NET 8 and .NET 10.
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

## Bound a batch by source bytes

Set `ReaderBatchOptions.MaxTotalBytes` to bound the combined source lengths accepted by a batch.
When set, reads run in input order so each native owner receives the remaining allowance;
`MaxDegreeOfParallelism` remains an upper bound. Per-file and registered default limits still apply.
Directory packages use their native bounded reader and snapshot length.

`ReadDocumentsDetailedAsync` and `ReadDocumentsAsCompletedAsync` report a source that does not fit
as a failed outcome and continue with the remaining sources. `ReadDocumentsAsync` throws instead.
The other operation budgets still span the complete batch; terminal safety limits retain their
existing behavior.

## Branch on failed inputs

Resilient batch and folder reads keep going when one input fails and report a stable code for it, so a caller can
treat a missing password differently from a damaged file without matching exception types or messages:

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.All;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddAllOfficeIMOHandlers().Build();
IReadOnlyList<ReaderDocumentReadOutcome> outcomes = await reader.ReadDocumentsDetailedAsync(
    Directory.EnumerateFiles("inbox"), new ReaderOptions { OpenPassword = "shared-secret" });

foreach (ReaderDocumentReadOutcome outcome in outcomes.Where(item => !item.Succeeded)) {
    string action = outcome.ErrorCode switch {
        ReaderErrorCodes.PasswordRequired or ReaderErrorCodes.InvalidPassword => "ask for the password",
        ReaderErrorCodes.InvalidContent => "quarantine",
        _ => "retry later"
    };
    Console.WriteLine($"{outcome.Path}: {outcome.ErrorCode} -> {action}");
}
```

`ReaderDocumentReadOutcome.ErrorCode` covers `ReadDocumentsDetailedAsync` and `ReadDocumentsAsCompletedAsync`.
Folder reads set `ReaderSourceDocument.ErrorCode` and `ReaderIngestFileResult.ErrorCode` on files skipped because
reading them failed; files skipped by a folder byte budget have no code. Single-document reads keep throwing; pass the
exception, and for a file its path and options, to `reader.ClassifyReadError(exception, path, options)` to get the same
code.

`ReaderErrorCodes` defines `password-required`, `invalid-password`, `unsupported-encryption`, `permission-denied`,
`unsupported-format`, `invalid-content`, `timeout`, `not-found`, `access-denied`, `io-error` and `read-failed`.
Codes are never renamed; treat a code you do not recognize as `read-failed`. The PDF, OpenDocument and email-store
handlers map their engines' password and encryption failures. The Office engines mark a password that did not verify,
which Reader reports as `invalid-password`; damaged encrypted content that fails its integrity check after the right
password stays `invalid-content`. Without `OpenPassword`, a failed Word, Excel or PowerPoint file that the encryption
detection below finds encrypted is reported as `password-required`. Legacy `.doc` files are not decrypted, so a wrong
password for one keeps the general code. A custom handler maps its own exceptions with
`ReaderHandlerRegistration.ClassifyReadError`. Caller errors such as invalid arguments or options
still throw, and resource and admission limits still end the whole operation.

## Detect encryption without opening

`Detect` reports `IsEncrypted` when bounded content inspection finds encryption, without opening or decrypting the
file:

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.All;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddAllOfficeIMOHandlers().Build();
ReaderDetectionResult detection = reader.Detect("quarterly.xlsx");
if (detection.IsEncrypted) {
    Console.WriteLine($"{detection.Kind} file is encrypted: {string.Join(", ", detection.Evidence)}");
}
```

It covers password-encrypted Office packages, legacy Word, Excel and PowerPoint files that mark themselves encrypted,
OpenDocument packages whose `META-INF/manifest.xml` declares encrypted entries, PDFs whose trailer has an encryption
dictionary, and ZIP archives with an encrypted entry. An `encryption:` evidence token names the signal. An encrypted
Office package stores its format inside the encrypted payload, so its `Kind` comes from a `.docx`, `.xlsx`, `.pptx` or
related extension. A PDF protected only by an owner password also reports `IsEncrypted` even though it opens without a
password.

`IsEncrypted = true` is a finding; `false` only means detection saw no encryption, not that the input is unencrypted.
Content is inspected in `PreferContent` mode, the standalone default, and in `ContentWhenUnknown` mode when the
extension is unknown; `ExtensionOnly` never inspects content. A non-seekable stream, or `InspectContainers = false`,
limits a ZIP archive to its first entry and skips the OpenDocument manifest; a non-seekable PDF is checked only within
the probed prefix. Inspection also stops at `MaxProbeBytes` and `MaxContainerEntries`.

## Reuse an immutable read policy

`OfficeDocumentReaderBuilder.WithDefaultOptions(options)` captures common read and folder settings.
Registration and `Build()` each snapshot the policy, including nested budgets, encoding and folder
selection. Changing the original options, the builder or the copy returned by `reader.DefaultOptions`
does not reconfigure an existing reader.

Omitting per-read options uses that profile across path, stream, bytes, incremental, asynchronous,
folder and batch routes. An explicit `ReaderOptions` object replaces the complete common profile;
it does not merge individual fields. Explicit `ReaderFolderOptions` similarly replace folder defaults.
Standalone detection retains its own `ReaderDetectionOptions` contract, and `DetectKind` remains an
extension lookup that does not open a file.

## Operation budgets and decoding

`ReaderOptions.CreateSafeIngestion()` sets a 64 MiB input limit and finite aggregate budgets for chunks,
characters, blocks, assets and nested content. Its values are a starting policy that applications can tighten:

```csharp
using OfficeIMO.Reader;

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

`LayoutTimeout` bounds the time one read spends laying out or rendering pages, such as Word page locations and DjVu
page images. It uses the same deadline as image export's `RenderTimeout`, so exceeding it throws
`OfficeImageExportTimeoutException`, a `TimeoutException`, while cancelling your own token still throws
`OperationCanceledException`:

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Word;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddWordHandler(new ReaderWordOptions { IncludePageLocations = true })
    .Build();
var options = ReaderOptions.CreateSafeIngestion();
options.LayoutTimeout = TimeSpan.FromSeconds(30);
try {
    OfficeDocumentReadResult document = reader.ReadDocument("long-report.docx", options);
} catch (TimeoutException) {
    Console.WriteLine("Page layout took longer than 30 seconds.");
}
```

Routes that do not lay out pages ignore the limit. Batch and folder reads report a timed-out input with the `timeout`
code and continue.

`MaxChars` is a best-effort projection target. Atomic Markdown blocks and tables can exceed it with a warning;
Reader preserves their complete content. Use `ResourceLimits.MaxChunkCharacters` for a terminal operation limit.
Folder and detailed path reads run the same configured processors as single-document reads.

Plain-text handlers recognize UTF-8, UTF-16 and UTF-32 BOMs. Without a BOM they use UTF-8 or the supplied
`TextEncoding`. Replacement decoding emits a warning; `ThrowOnInvalidTextBytes` rejects malformed bytes.
BOMs override `TextEncoding`. These options apply to the plain-text and source-code handlers; other formats use their own
codec policies. Chunk boundaries preserve Unicode surrogate pairs, and line breaks normalize to LF.

## Text and metadata projection

Supporting handlers can read text and metadata without building secondary Markdown, asset, visual,
form or OCR projections:

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Word;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddWordHandler().Build();

var options = ReaderOptions.CreateSafeIngestion();
options.Projection = ReaderProjection.TextAndMetadata;
OfficeDocumentReadResult document = reader.ReadDocument("Policy.docx", options);
```

The default is `ReaderProjection.Full`. Check each handler's `SupportsTextAndMetadata` capability
and `TextAndMetadataExtensions` before requesting the lean projection. An empty extension list means
no extension restriction for a supporting handler; a nonempty list identifies its qualified subset.
Unsupported handlers and formats reject the request before parsing rather than silently running a full read.
Several legacy Office imports reconstruct pictures as part of loading and therefore remain Full-only.

The lean result retains text, source locations, metadata and configured native inspection evidence.
Its `CapabilitiesUsed` includes `officeimo.reader.projection.text-and-metadata`; an empty asset collection
in that profile means asset projection was not requested, rather than proving the source contains no images.
Options are snapshotted, and nested reads inherit the lean operation profile. Core does not reconstruct
unrequested Markdown when refreshing a text processor's changes.

This projection avoids secondary work; it is not a hard process-memory limit. Keep native parser,
package expansion and input limits alongside aggregate output budgets. Format owners document
any source structures that must still be decoded or validated during loading.

## Text projection

`ReaderProjection.Text` returns only text chunks. `GetText()` gives the same text as the
text-and-metadata projection, but the result carries no blocks, tables, pages, links,
assets, metadata, Markdown or nested results; it keeps the source, diagnostics and
`CapabilitiesUsed`, which includes `officeimo.reader.projection.text`:

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Word;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddWordHandler().Build();

var options = ReaderOptions.CreateSafeIngestion();
options.Projection = ReaderProjection.Text;
string text = reader.ReadDocument("Policy.docx", options).GetText();
```

Every handler that supports the text-and-metadata projection supports this one, and the
same handlers reject it. A handler may read the text directly without building its
document model; otherwise Reader reduces the text-and-metadata result to one chunk per
text block, fallback chunk and table row that `GetText()` writes. Each chunk keeps its
source location, with its position in the text as `LogicalOrder`. The result has the
same shape either way. Nested reads inherit the text profile.

## Nested results and qualified capabilities

ZIP, email and PDF attachment reads keep each delegated source's rich result in `NestedDocuments`, alongside the existing
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

Capability manifest version 7 includes incremental route flags, per-extension `FormatQualifications`, and lean-projection support.
A qualification records a format ID, extraction maturity, profile, preservation, limitations and evidence
references. Word, Excel and modern PowerPoint reuse their owner's format IDs. Other undeclared profiles
remain `Unqualified`; an available handler is not a claim of complete format preservation. Applications can
supply immutable `ReaderFormatQualification` values when registering a custom handler.

For model-specific token budgets, pass a `ReaderDelegateTokenCounter` to `ReaderHierarchicalChunkingOptions`:

```csharp
using OfficeIMO.Reader;

var tokenizer = new WhitespaceTokenizer();
var chunking = new ReaderHierarchicalChunkingOptions {
    MaxTokens = 800,
    TokenCounter = new ReaderDelegateTokenCounter("my-model:vocabulary-version", tokenizer.CountTokens)
};

// Stand-in for your model's tokenizer.
sealed class WhitespaceTokenizer {
    public int CountTokens(string text) => text.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries).Length;
}
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
`tables.ExportTables(cancellationToken: token)` for CSV, Markdown and JSON together, or use
`table.ToCsv(token)`, `table.ToMarkdownTable(token)` and `table.ToJson(indented: true, cancellationToken: token)` separately.

- `ReaderOptions` and format-neutral input/processing limits
- `ReaderChunk` and the schema-versioned `OfficeDocumentReadResult`
- tables, pages, visuals, assets, links, forms, metadata, and diagnostics
- sync/async path, stream, byte-array, folder, and batch ingestion
- deterministic capability manifests with `OfficeIMO` versus `Custom` handler origins
- bounded nested-content delegation between configured handlers

Public namespaces remain `OfficeIMO.Reader`; the `.Core` name describes package and assembly ownership.
