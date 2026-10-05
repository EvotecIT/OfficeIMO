# OfficeIMO.Workflows

`OfficeIMO.Workflows` is the reusable local orchestration layer for OfficeIMO document jobs. It composes the existing first-party conversion, PDF, and provenance APIs behind typed requests, bounded execution, cooperative cancellation, collision policies, atomic output publication, and post-write validation.

The package does not add a second document or PDF engine. Desktop applications, command-line tools, and services can share this workflow contract while keeping their user-interface and hosting code thin.

## Opt-in Apple conversion

[OfficeIMO.Workflows.IWork](../OfficeIMO.Workflows.IWork/README.md) supplies Pages-to-Word, Numbers-to-Excel, and Keynote-to-PowerPoint routes. `IWorkWorkflow.CreateRunner()` shares this runner's source capture, limits, destination reopen validation, and publication contract. `IOfficeWorkflowRunner.ConversionRoutes` exposes the configured executable routes; the static `OfficeWorkflowCatalog` describes built-in executability. The default workflow package remains independent of iWork.

`OfficeWorkflowConversionRegistration` adds an implementation of an existing canonical route to a runner. It cannot replace built-in owners. Opt-in converters accept captured ZIP/file streams, write to a bounded caller-owned output stream, and return immutable `OfficeWorkflowConversionEvidence`. Current opt-in destination formats are DOCX, XLSX, and PPTX. `OfficeWorkflowResult.ConversionEvidence` retains fidelity categories and compact source facts; successful reopen does not establish visual equivalence.

Provider directory packages use `OfficeWorkflowRequest.InputDirectoryPackage` with a registered directory-package converter. The shared runner preserves member layout in bounded private staging and verifies original provider membership and content before publication. The host supplies permission-aware root identity and output-separation checks through `OfficeWorkflowDirectoryPackageInput.SourcePublicationGuard`. The selected filename determines routing; an explicit output is required.

## Single conversions and file batches

`OfficeWorkflowRunner` executes its configured `ConversionRoutes`. Ordinary directory and selected-file batches use those same routes, including registered adapters, profiles, renderer options, diagnostics and publication policies. Registered adapters use the options captured when the runner was created. PDF export covers DOC, DOCX, TXT, XLSX, PPTX, HTML, Markdown and RTF. Other built-in targets use the existing PDF-to-DOCX/XLSX/PPTX/HTML routes. Unsupported or filtered files produce skipped outcomes; their count is separate from selected conversions.

```csharp
using OfficeIMO.Pdf;
using OfficeIMO.Workflows;

var options = new OfficeWorkflowConversionOptions {
    PlainText = new PdfPlainTextOptions { TabSize = 4 }
};
var runner = new OfficeWorkflowRunner();
var single = await runner.RunAsync(new OfficeWorkflowRequest {
    Operation = OfficeWorkflowOperation.Convert,
    InputPath = "report.txt", OutputPath = "report.pdf", ConversionRouteId = "txt-pdf", ConversionOptions = options
}, cancellationToken: cancellationToken);

OfficeConversionBatchResult batch = await OfficeWorkflow.ConvertDirectory("Documents")
    .ToDirectory("PDF")
    .WithConversionOptions(options)
    .WithConcurrency(2)
    .RunAsync(runner, cancellationToken: cancellationToken);

// Selected files use the same batch API; the target need not be PDF.
var html = await OfficeWorkflow.ConvertFiles("report.pdf", "appendix.pdf")
    .ToDirectory("HTML", ".html")
    .RunAsync(runner, cancellationToken: cancellationToken);
```

`Word`, `Excel`, `PowerPoint`, `Html`, `Markdown`, `Rtf` and `PlainText` accept their owning adapter's typed options. A mixed batch selects the settings applicable to each route. Explicit renderer options take precedence over the cross-format `OutputProfile`. For ambiguous source extensions, use `.Via("html-pdf")` or `ConversionRouteId` on the request; TXT otherwise remains literal text. Encrypted Office inputs use `ConversionOptions.SourcePassword`; PDF inputs use `PdfPassword`. Passwords remain runtime inputs and are not stored in checkpoints.

Ordinary batches support the existing `Fail`, `Rename` and `Replace` conflict policies. Directory discovery is incremental and skips filesystem links. Outputs retain the full relative source filename plus the target extension, so `report.doc` and `report.docx` have distinct PDF names. Explicit files retain relative paths when `InputDirectory` supplies their common root; otherwise they use their filenames, and destination collisions follow the selected policy.

## Book publishing

`BookManuscriptImporter` composes the owning Word, Markdown, HTML and EPUB libraries.
It imports `.docx`, `.md`, `.markdown`, `.html` and `.htm` manuscripts into reflowable
books, retaining each conversion stage's fidelity categories. Word uses its static
final revision view: comments are omitted, fields use their stored visible results,
and live controls are outside the book contract. Notes and supported semantic content
remain in the publication. Markdown front matter supplies title, language and author.

```csharp
using OfficeIMO.Epub;
using OfficeIMO.Workflows;

var imported = await BookManuscriptImporter.ImportFileAsync("manuscript.md",
    new EpubManuscriptOptions { ChapterHeadingLevel = 1 });
BookProject project = BookProject.FromImport(imported);
imported.Report.RequireNoLoss();
project.RenameChapter(0, "Opening chapter");
project.SetStylesheet(EpubTypography.CreateStylesheet(EpubTypographyProfile.Prose));
await File.WriteAllBytesAsync("book.oibook", project.ToProjectBytes());
BookProject reopened = BookProject.LoadProject(await File.ReadAllBytesAsync("book.oibook"));
await File.WriteAllBytesAsync("book.epub", reopened.Export().Bytes);
```

The file route allows at most 64 MiB of manuscript input and resolves assets only
inside the manuscript's physical parent directory. It rejects executable Word
package parts. `ImportBytesAsync` consumes a host-owned snapshot and does not read
files implicitly; a host may supply a permission-aware resource resolver and base URI.
Typed `ImportWordAsync` and `ImportMarkdownAsync` reuse an already loaded source.

`BookProject` owns validated edits: metadata, chapter insertion/removal/reordering,
chapter titles and XHTML bodies, resource renaming with reference repair, a project stylesheet, and cover selection.
`RenameResource(manifestId, containerPath)` delegates to the EPUB owner and retains the same undo/redo behavior as other project edits. Its resource-inspection limits are described in the EPUB README.

`SplitChapter(manifestId, boundaryId, newManifestId, newContainerPath, title)` delegates the atomic chapter split to the EPUB owner. The resulting content, navigation and reading-order changes participate in project undo/redo and persistence. See the EPUB README for supported boundaries and reference-repair limits.

`MergeChapters(firstManifestId, secondManifestId, boundaryId)` combines consecutive compatible chapters through the same owner and undoable transaction. Both navigation entries survive, and the second chapter's links target retained content or its new boundary. Conflicting styles, identifiers and metadata require explicit resolution; see the EPUB README for the merge contract.
`ApplyEdits` commits a complete editor draft atomically. Invalid or cancelled edits
retain the previous publication. Deleting a linked chapter requires repairing its
remaining links first. A blank creator retains the current creator. Package edits
have one bounded session-only undo/redo step. Named revisions are saved separately.
`PreviewChapter` renders through `OfficeIMO.Epub.Image` using retained package assets.
It selects the requested spine position and fails if that chapter was omitted by the
bounded reading policy. Navigation edits also reject incomplete reader projections,
retaining the complete publication when item, depth, or XML size limits are reached.

Capture editorial milestones explicitly with `CreateRevision(name)`. `Revisions`
returns immutable descriptors with an ID, name, UTC timestamp, SHA-256 hash and byte
count. `RestoreRevision(id)` restores publication content as an undoable edit;
`RemoveRevision(id)` removes only that snapshot. Import diagnostics and acceptance
remain project-wide. Direct edits through `Publication` are included when capturing
a revision, but are not automatically recorded as history.

```csharp
var baseline = project.CreateRevision("Before copyediting");
project.SetMetadata("Revised title", "en", "Author");
byte[] saved = project.ToProjectBytes();
var restored = BookProject.LoadProject(saved);
restored.RestoreRevision(baseline.Id);
restored.Undo(); // Return to the revised title.
```

A project retains up to 100 named revisions and 128 MiB of combined revision EPUB
bytes, in addition to its current publication (up to 128 MiB). Capture rejects an
exhausted bound without evicting existing revisions. Version-2 projects retain these
snapshots; version-1 projects remain readable. Session undo/redo is not persisted.
Revision hashes detect inconsistent stored content; they are not digital signatures.

The `.oibook` container stores `publication.epub`, named revision EPUBs and a versioned review record, with
physical ZIP validation and byte/count limits. Loading never extracts files.
Projects may retain non-fatal review findings until the author acknowledges them;
failure diagnostics cannot be accepted as export-ready. Every EPUB export still runs
the native writer's validation. Project review records are user-owned state, not an
authenticity certificate. Project instances are mutable and not thread-safe.
Hosts own destination permissions, conflict handling and safe publication; Studio
uses its existing verified storage owner for those operations.

## Optional checkpoints

Add `.WithCheckpoint("PDF-State")` to the builder, or set `CheckpointDirectory` on `OfficeConversionBatchRequest`, for restartable execution. Source, output and checkpoint trees must be separate local folders. For selected HTML files and Markdown files with local resources enabled, output and checkpoint folders must also be outside each file's resource tree, including an explicit Markdown `BaseDirectory`. Checkpoint jobs require `Fail`: recorded completed artifacts are immutable and verified by source, rendering-settings, local-resource and output hashes before reuse.

Checkpoints support built-in routes. Registered adapters require an ordinary batch because their captured runtime configuration cannot be fingerprinted. An explicit registered route with checkpoints is rejected before execution; a registered input discovered in a mixed checkpoint batch reports a failed item without publishing it.

Before publication, the runner flushes validated staged output and records its hash and staging identity. Restart can finish that recorded move or verify an output moved before the final receipt was written. Changed completed sources or settings, altered/missing outputs and outputs without a bound receipt fail the item for inspection. `RetryFailed` permits retrying recorded failures, including corrected failed inputs. Completed files and recorded pending publications survive cancellation. An interruption before publication intent is recorded can leave a hidden staging file; inspect it before removing it.

Checkpoint reuse verifies a **recorded artifact**; it does not rerender it or promise that a newer renderer would produce identical bytes. Compatible engine updates do not invalidate completed receipts. The checkpoint schema, host, source/output roots, target and per-item rendering inputs define compatibility. Execution concurrency, selection and byte/file budgets may change; the new budgets still apply when verifying artifacts. New source files are discovered on each run.

HTML and enabled local Markdown resources are conservatively fingerprinted within the source root, with at most 256 regular files and an aggregate input-byte budget. Every filename is included because Markdown identifies image formats from their bytes. Resource trees cannot contain links; checkpointed Markdown also requires `RestrictLocalImagesToBaseDirectory`. A changed CSS, image or font invalidates reuse. Checkpoints exclude remote resources and runtime resource, text-shaping or cryptography callbacks because their output cannot be identified from captured settings; use an ordinary batch or the native adapter for those cases. Workflow HTML resource resolution remains scoped to the source; custom HTML resolvers belong to the native adapter.

Per-document defaults are 64 MiB input and 256 MiB output; concurrency accepts 1–32. `MaximumFiles` bounds discovered files, including skipped files, and defaults to one million. Discovery beyond this bound stops the run while preserving completed output. These are configurable resource bounds, not throughput guarantees. TXT defaults to strict BOM-aware decoding, literal markup, tab expansion and bounded wrapping; `PlainText` carries its encoding and layout limits. Legacy DOC import loss blocks output unless `LegacyDocLossPolicy = OfficeConversionLossPolicy.Allow` accepts reported reductions.

`IProgress<OfficeConversionBatchItemResult>` callbacks can arrive concurrently; consume or stream them without retaining a whole inventory. The result contains bounded counts. Checkpoints retain up to 32 non-information diagnostics and report truncation. A host can supply `publicationGuard` to protect output and checkpoint destinations. Conversion completion retains each adapter's fidelity limits and does not prove exact Microsoft Office pagination.

Studio exposes **Convert → Batch PDF export**. The CLI uses `officeimo workflow batch`; PSWriteOffice uses `Export-OfficeDocumentPdf -InputDirectory ... -OutputDirectory ...` or selected file pipelines on PowerShell 7.4 or newer.

## Email evidence and conversation dossiers

`EmailEvidenceWorkflow` produces a portable ZIP containing `report.html`, `report.md`, `manifest.json`
and an optional `report.pdf`. Reports include From/To/Cc, sent and received dates, attachment indexes,
source fingerprints, protection classification, diagnostics and explicit body clipping. The body is
semantic text from `OfficeIMO.Email.Html`, escaped for display; original formatting and embedded images
are omitted. Attachment payloads and original messages stay outside the ZIP. The workflow reads local
content without network access, signature verification, decryption or certificate discovery.

```csharp
var evidence = EmailEvidenceWorkflow.Create("message.eml");
File.WriteAllBytes("message-evidence.zip", evidence.ToZipBytes());

using var mailbox = OfficeIMO.Email.Store.EmailStoreSession.Open("archive.pst");
var selected = mailbox.EnumerateItems().First().Key;
var dossier = EmailEvidenceWorkflow.CreateConversation(mailbox, selected,
    new EmailEvidenceOptions { MaxItemsScanned = 10_000, MaxMessages = 100 });
File.WriteAllBytes("conversation.zip", dossier.ToZipBytes());
```

Conversation selection reuses the existing graph. Messages are chronological; thread links retain their
evidence and heuristic status, and missing or ambiguous parents remain visible. `GraphComplete` reports
the graph owner's coverage. Each message is projected under the body/report bounds before the next body
is read; eager store formats retain their own bounded opening behavior. Embedded attachments are classified
from available MAPI metadata even when their nested payload is not read. A file fingerprint hashes the same open source before and after parsing;
a store fingerprint uses the store's durable source contract, including its composite directory hash.
Hashes identify source bytes and resident attachment payloads; they do not certify message authenticity.
Deferred attachment streams are not opened for hashing. Report fields use bounded display values,
diagnostics retain a sample of up to 500 entries with the total count, and PDF conversion diagnostics
are included in the manifest. `EmailEvidenceOptions` controls input, body, graph, report, page and output
bounds. Outputs are created in memory; applications choose and authorize their publication destination.

## Invoice inspection, conversion and rendering

`OfficeInvoiceBufferWorkflow` composes the typed invoice model, optional standards
validator and PDF adapter. It accepts captured XML bytes and returns an operation
report without reading paths, fetching invoice links or publishing files:

```csharp
using OfficeIMO.Invoicing;
using OfficeIMO.Workflows;

var target = new InvoiceXmlOptions(
    InvoiceSpecificationRelease.En16931_1_3_16,
    InvoiceSyntax.Ubl,
    InvoiceProfile.En16931);
var request = new OfficeInvoiceWorkflowRequest(
    File.ReadAllBytes("invoice.xml"),
    OfficeInvoiceWorkflowOperation.Convert,
    target,
    inputName: "invoice.xml");
var result = await OfficeInvoiceBufferWorkflow.RunAsync(request);
foreach (var diagnostic in result.Diagnostics)
    Console.WriteLine($"{diagnostic.Location}: {diagnostic.Message}");
if (result.Succeeded)
    File.WriteAllBytes("converted.xml", result.ToOutputBytes()!);
```

Inspection completion does not establish validity. Check `ModelValidation`,
`Source.HasCompleteMapping` and target diagnostics separately. Recognized
MINIMUM and BASIC WL inputs receive aggregate model checks without inventing
invoice lines. Conversion and rendering block unmapped source data and
unsupported target fields; explicit lower-profile projection returns each
intentional reduction as a warning.

Select `Validate`, `RenderPresentationPdf` or `RenderHybridPdf` for the other
operations. Rendering requires an explicit CII contract. Pass `PdfOptions` with
the fonts your content needs and `InvoicePdfLayoutOptions` for appearance and
resource limits. The returned output XML is the same captured invoice used for
the visible PDF; hybrid output embeds those exact bytes.

For bounded header replacements that retain XML extensions, create a request with
`OfficeInvoiceWorkflowRequest.ForSourceEdit(xml, new InvoiceSourceEdits(number:
"INV-002"))`. The file equivalent is
`OfficeInvoiceFileWorkflowRequest.ForSourceEdit("invoice.xml", "edited.xml",
edits)`, which uses the same output preflight and atomic publication contract.
`EditSource` retains the original syntax/profile and accepts no conversion target.
Its `Succeeded` status means every requested edit completed; model and mapping
findings can still contain errors. `Source` and `ModelValidation` describe the
edited XML, or remain unavailable when its retained data exceeds the semantic
mapper. Passing an explicit validation release and validator requires the exact
edited XML to pass schema and business rules before any artifact is returned.
See the [source-editing contract](../OfficeIMO.Invoicing/README.md#read-and-edit-safely)
for supported fields, representations and bounds.

Standards validation requires both `validationRelease` and a configured
`InvoiceValidator` supplied to `RunAsync`. A requested validator that is missing,
or a schema/business-rule stage that does not pass, blocks output. Otherwise
`SchemaStatus` and `BusinessRulesStatus` explicitly report `NotRun`. For writing
operations, `StandardsValidation.Sha256` identifies the output XML validated
before artifact generation; it does not certify the PDF's conformance.

`RunBatchAsync` preflights requests before executing them, preserves input order,
and observes cancellation between bounded owner operations. Defaults are 256
requests, 64 MiB of combined XML input and 64 MiB of retained output artifacts.
`ContinueOnFailure` controls whether subsequent items run. An item exceeding the
output budget returns diagnostics and no artifact bytes. Cancellation throws
`OperationCanceledException`; hosts remain responsible for collision policies
and safe output publication.

`OfficeInvoiceFileWorkflow` provides that local-file adapter. Its immutable
`OfficeInvoiceFileWorkflowRequest` captures paths and the same target and render
settings. It preflights all inputs and destinations, applies the combined batch
budgets, then creates each successful artifact through the shared atomic writer.
Existing destinations and colliding batch outputs are rejected. It never
overwrites inputs or existing files; batch publication is per item rather than a
transaction. `OfficeInvoiceFileWorkflowResult` keeps publication errors separate
from the model and standards evidence in `Workflow`.

For a desktop host with local or provider-backed storage, use
`OfficeWorkflowRunner.RunInvoiceAsync` with `OfficeInvoiceStorageWorkflowRequest`:

```csharp
var storageResult = await new OfficeWorkflowRunner().RunInvoiceAsync(new() {
    InputPath = "invoice.xml",
    Operation = OfficeInvoiceWorkflowOperation.EditSource,
    SourceEdits = new InvoiceSourceEdits(number: "INV-002"),
    OutputPath = "invoice.edited.xml",
    ConflictPolicy = OfficeWorkflowConflictPolicy.Rename
});
```

The adapter captures at most 16 MiB of input, clones render settings before
acquisition, and verifies source contents and physical identity again before
publication. Local output supports fail, numbered-copy and atomic replacement
policies. It protects source aliases and asks the supplied publication guard
about the final destination. Provider inputs use reopenable
`OfficeWorkflowStreamInput`; provider output uses `OfficeWorkflowStreamOutput`
with explicit `Replace` after the host obtains direct-write consent. A durable,
hash-verified XML or PDF recovery copy precedes provider creation/writing. Failed
or unverified provider publication returns `Unconfirmed` with retained recovery;
it cannot promise atomic replacement or rollback. Read `Workflow` for invoice
evidence and `Status`, `Diagnostics` and `Recovery` for storage outcomes. The
default retained output limit is 64 MiB. Cancellation before publication returns
`Cancelled` and removes temporary staging.

## Project reports and table exchange

`ProjectReportWorkflow` exports a calculated Project view through the existing document owners:

```csharp
using OfficeIMO.Project;
using OfficeIMO.Workflows;

using var project = ProjectDocument.Load("delivery.xml");
var schedule = project.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true });
schedule.Report.ThrowIfErrors();
var view = project.CreateView(schedule, new ProjectViewOptions {
    Kind = ProjectViewKind.TaskUsage,
    Timescale = ProjectViewTimescale.Week
});
File.WriteAllBytes("delivery.pdf", ProjectReportWorkflow.ToPdf(view));
File.WriteAllText("delivery.html", ProjectReportWorkflow.ToHtml(view));
using var workbook = ProjectReportWorkflow.CreateExcel(view);
workbook.Save("delivery.xlsx");
```

`ToSvg` and `ToPng` return one result per page. Supply an `OfficeRenderingProfile` with the required fonts for explicit Unicode coverage. `CreateWord` and `CreatePowerPoint` include chart images followed by editable data tables. `CreateExcel` separates report values, usage, status/groups, baseline dates, and dependencies into editable worksheets. These exports do not reconstruct Microsoft Project's saved views or native styles.

Project PNG exports default to 300 DPI, rendered from the drawing at the requested resolution. Choose a shared quality preset for screen images:

```csharp
using OfficeIMO.Drawing;

var images = ProjectReportWorkflow.Images(view)
    .WithQuality(OfficeImageExportQuality.Screen)
    .As(OfficeImageExportFormat.Png)
    .Export();
```

The presets are `Preview` (96 DPI), `Screen` (192 DPI), and `Print` (300 DPI). `ExportImages` returns encoded bytes, pixel dimensions, density, and diagnostics; its consumer overload streams results under shared batch limits. Supply `ProjectImageExportOptions` to select fonts, explicit density, pixel limits, or a rendering deadline. Oversized Project images fail by default; choose `RasterOverflowBehavior.ReduceScale` only when reduced detail is acceptable. Use SVG for zoomable vector text and geometry. Enlarging an existing PNG beyond its pixel dimensions still magnifies its pixels.

For consistent typography, register both regular and bold TrueType faces in the rendering profile used for measurement and output. A regular face alone requires synthesized bold text; font substitution can change line wrapping. `ProjectOfficeReportOptions` selects chart images, data tables, and chart image quality for Word and PowerPoint. Set `IncludeCharts = false` for editable-table reports. A Table view always retains its primary editable content, including when only charts are selected.

Word tables repeat their headers and flow across pages. PowerPoint uses measured row heights to keep complete rows on each slide. Excel retains numeric and date cells, freezes the header row, and prints narrow reports in portrait and wider usage tables in landscape, with a report title and page numbers. Tables wider than eight columns print at full scale across pages; usage sheets repeat UID and name columns on horizontal continuations. Native Office exports retain fixed page dimensions; portable drawing exports can trim unused page height through `ProjectViewOptions.FitPageHeightToContent`.

`ProjectDataWorkflow` transports the Project owner's mapped tables:

```csharp
var projection = project.ExportTables(allowLossyProjection: true);
foreach (string notice in projection.Notices)
    Console.WriteLine(notice);
using var transfer = ProjectDataWorkflow.CreateExcel(projection);
transfer.Save("project-data.xlsx");
foreach (var table in projection.Tables)
    ProjectDataWorkflow.CreateCsv(table.Table).Save(table.Kind + ".csv");
```

Loss permission is explicit because tables omit dependencies, native presentation, and other semantics outside the selected exchange fields. `ReadExcel` requires a bounded worksheet rectangle; formula/error cells and numbers that cannot be represented exactly are rejected. `ReadCsv` uses the CSV owner's parsing and quoting rules. Wrap the resulting `ProjectDataTable` in a `ProjectMappedTable` with explicit field mappings before calling `ProjectDocument.ImportTables`. Table import creates a new project and validates identity, references, units, and conflict policy. See [Project support](../OfficeIMO.Project/SUPPORT.md#portable-reports-and-mapped-data-exchange) for the full boundary.

## Review OCR before publication

`MakePdfSearchableAsync` captures recognition evidence before writing an output. Set `PdfSearchableWorkflowRequest.ReviewAsync` to choose eligible words, or `ReviewCorrectionsAsync` to return original eligible word instances mapped to their reviewed text. Choose one callback. An empty review selection deliberately preserves the source copy; a recognition result with no eligible words and no native source text fails before publication.

```csharp
request.ReviewCorrectionsAsync = (review, token) => {
    token.ThrowIfCancellationRequested();
    IReadOnlyDictionary<OfficeIMO.Pdf.Ocr.PdfRecognizedWord, string> reviewed = review.Ocr.Pages
        .SelectMany(page => page.Words).ToDictionary(word => word, word => word.Text);
    return Task.FromResult(reviewed);
};
```

A host review interface can edit dictionary values and exclude entries before returning them. Correction eligibility, text limits, source-identity checks, output conflicts, and publication guards apply to local files, provider outputs, and OCR sessions. Workflow diagnostics retain recognition warnings and page numbers. A nonrecoverable provider error prevents publication, and image recognition with no usable text does not create an empty success artifact. Successful publication means the reviewed artifact was saved; it does not certify recognition accuracy.

OfficeIMO Studio shows the source region, original recognition, editable replacement, confidence, and inclusion choice. Corrections persist across page navigation. **Next uncertain word** navigates low-confidence and sub-90% words without making rejected words eligible. The selected text can be extracted without creating a PDF or saved as a searchable layer; cancellation preserves the existing destination.

## Reference from source

When working from an OfficeIMO source checkout, reference the workflow project directly:

```xml
<ProjectReference Include="..\OfficeIMO.Workflows\OfficeIMO.Workflows.csproj" />
```

## Inspect concealed text in raster images

`OfficeRasterContentSafety` combines a caller-supplied `IOcrEngine` with decoded pixel evidence. It accepts one static raster image, normalizes it to a metadata-free PNG, and assesses every OCR line, word, or character span that includes bounded pixel or normalized geometry. Image metadata and unbounded provider text do not become visibility findings.

```csharp
using OfficeIMO.ContentSafety;
using OfficeIMO.Ocr;
using OfficeIMO.Workflows;

IOcrEngine engine = GetConfiguredOcrEngine();
byte[] source = File.ReadAllBytes("review.png");

var options = new OfficeRasterContentSafetyOptions {
    EnableOpaqueRectangleRedaction = true,
    MinimumOcrConfidenceForRedaction = 0.9
};

OfficeContentSafetyReport report = await OfficeRasterContentSafety.InspectAsync(
    source,
    engine,
    options,
    cancellationToken);

string[] approvedIds = report.Findings
    .Where(finding => finding.CleanupCapability == OfficeContentCleanupCapability.RedactRegion)
    .Select(finding => finding.Id)
    .ToArray();

OfficeContentCleanupResult result = await OfficeRasterContentSafety.RedactSelectedContentAsync(
    source,
    engine,
    new OfficeContentCleanupSelection(approvedIds),
    options,
    cancellationToken);
File.WriteAllBytes("review-redacted.png", result.Output);
```

The inspector reports nearly transparent, tiny, and low-contrast OCR regions using bounded geometry and conservative pixel evidence. Aggregate OCR text must be fully represented by accepted bounded spans after verified hierarchical duplicates are removed, and malformed Unicode is rejected before finding identities are derived. Redaction is deliberately disabled by default. When enabled, only sufficiently confident, currently matching findings can be selected; the workflow covers their bounded regions with the configured opaque color, emits a single-frame PNG derivative, reopens it, verifies every output pixel, and reruns OCR inspection under the same captured engine identity and capabilities. A selected region, including its configured padding, must not overlap any independently bounded recognized span that was not also selected. When a provider emits an aggregate line or word together with finer spans carrying the same line identity, the finer spans own overlap validation only when they fully reproduce the aggregate text. Cumulative limits bound both region-pixel work and OCR-region intersection comparisons. This is destructive rectangular coverage, not semantic image editing. Multi-frame images, unsupported geometry, oversized provider output, input or analysis work, provider errors or non-recoverable diagnostics, color-rendering metadata that is neither canonical sRGB nor normalized by the managed decoder, unsupported embedded orientation, and outputs where OCR still recognizes leaf text in a changed region all fail closed.

## Save a prepared scan copy

`ScanCleanup` creates a separate PDF containing the prepared page pixels. It uses the same page selection, region crop, perspective correction, and tonal settings as `OfficeIMO.Pdf.Ocr` preview. The source is protected from replacement. Native text, forms, links, signatures, and attachments are omitted from the raster copy, so callers must acknowledge that output contract.

```csharp
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;
using OfficeIMO.Workflows;

OfficeWorkflowResult result = await new OfficeWorkflowRunner().RunAsync(new() {
    Operation = OfficeWorkflowOperation.ScanCleanup,
    InputPath = "scan.pdf",
    OutputPath = "prepared-scan.pdf",
    ScanCleanup = new() {
        AcknowledgeRasterOutput = true,
        Preparation = new() {
            Dpi = 200,
            ReadOptions = new() { PageSelection = PdfPageSelection.From(1) },
            ScanProcessing = new() { Deskew = false, StraightenDegrees = 2, Gamma = 1.1 }
        }
    }
});
```

Set `ExpectedSourceSha256` to the SHA-256 hex digest of a reviewed snapshot to reject a source that changed before export. Provider inputs and destinations use the same snapshot, confirmation, recovery, and publication guards as other workflows. To retain the visible source and add searchable text, use the searchable OCR workflow instead.

## Optimize embedded Word images

```csharp
using OfficeIMO.Drawing;
using OfficeIMO.Word;
using OfficeIMO.Workflows;

var runner = new OfficeWorkflowRunner();
OfficeWorkflowResult result = await runner.RunAsync(new OfficeWorkflowRequest {
    Operation = OfficeWorkflowOperation.OptimizeWordImages,
    InputPath = "input.docx",
    OutputPath = "optimized.docx", // .doc or .pdf also selects that output format
    ConflictPolicy = OfficeWorkflowConflictPolicy.Fail,
    WordImageOptimization = new WordImageOptimizationOptions {
        Mode = OfficeImageOptimizationMode.DownsampleAndRecompress,
        TargetDpi = 144,
        JpegQuality = 85
    }
});
```

The runner snapshots the source, optimizes through `OfficeIMO.Word`, reopens the staged output through its format owner, and publishes a separate copy atomically. Source replacement and publication over any batch source are refused. `AnalyzeWordImages` returns per-media diagnostics without publishing a file. Diagnostics include dimensions, formats, candidate metadata removal and whether each change was applied. Metadata removal produces a warning, and the inventory exposes `requiredStagedBytes` for budgeting. `RunBatchAsync` accepts up to 250 requests, snapshots options before execution, and publishes each item independently.

DOCX and supported legacy DOC inputs can produce DOCX, native DOC, or PDF. Incomplete legacy projections block output; analysis warns that its inventory covers only projected pictures. The native DOC writer preflights destination support. Word reports encoded-media savings; `InputBytes` and `OutputBytes` measure actual files. PDF generation after Word optimization retains its default image policy to avoid a second JPEG quality reduction. [Word image options and preservation rules](../OfficeIMO.Word/README.md#images) apply to every host.

## Convert a document

The executable conversion routes run in process through the OfficeIMO format and
rendering packages. They do not launch an external office suite or document
converter. Independent producer files and compatibility checks belong to
validation; they are not prerequisites for running these conversions.


```csharp
using OfficeIMO.Workflows;

var runner = new OfficeWorkflowRunner();
OfficeWorkflowResult result = await runner.RunAsync(new OfficeWorkflowRequest {
    Operation = OfficeWorkflowOperation.Convert,
    ConversionRouteId = "docx-pdf",
    InputPath = "report.docx",
    OutputPath = "report.pdf",
    ConflictPolicy = OfficeWorkflowConflictPolicy.Replace
});

if (!result.Succeeded) {
    throw new InvalidOperationException(result.Summary);
}

Console.WriteLine(result.OutputPath);
```

The same request has a fluent form. When a source/target extension pair maps to
one locally executable route, the builder infers it; use `Via(routeId)` for an
ambiguous pair or to pin a specific route contract.

```csharp
OfficeWorkflowResult result = await OfficeWorkflow
    .Convert("report.docx")
    .To("report.pdf")
    .WithProfile(OfficeWorkflowOutputProfile.PrintReady)
    .OnConflict(OfficeWorkflowConflictPolicy.Replace)
    .RunAsync(cancellationToken: cancellationToken);
```

`OfficeWorkflowCatalog.Routes` projects the complete canonical first-party
conversion catalog for discovery. `ExecutableRoutes` is the subset this local
runner can invoke. Route metadata includes accepted extensions, the owning
package, representative API and result contract, fidelity and support evidence,
known limits, browser and agent availability, and `CanExecute`.

Use `ConversionOptions` (or the fluent `WithConversionOptions` method) for settings
specific to a route. PDF input routes accept `PageRanges`. PDF-to-Word and
PDF-to-PowerPoint expose editable or visual import modes; visual pages accept
`RasterDpi`. Excel-to-PDF supports worksheet-canvas or flowing-table layout, and
PDF-to-HTML supports semantic or positioned HTML. PDF output routes can request
verified lossless compression with `CompressPdfOutput`. The runner rejects options
and output profiles unsupported by the selected route.

```csharp
OfficeWorkflowResult result = await OfficeWorkflow.Convert("source.pdf")
    .To("visual-pages.docx")
    .WithConversionOptions(new OfficeWorkflowConversionOptions {
        PageRanges = "1-3,5",
        WordMode = OfficeIMO.Word.Pdf.PdfWordImportMode.VisualPages,
        RasterDpi = 144
    })
    .RunAsync(cancellationToken: cancellationToken);
```

`OfficeWorkflowRunner.PreviewDocument(bytes, extension, cancellationToken)` creates
an in-memory sample of the first three pages of a PDF, DOCX, XLSX, PPTX, or HTML
artifact. The sample includes rendering diagnostics and uses bounded input and
image sizes. HTML previews do not load external or sibling resources. Previewing
is a review aid; inspect the full saved document for whole-document fidelity.

`RunAsync` also exposes PDF inspection, comparison, optimization, repair planning, repair, and sanitization through typed operations. `ExportPdfPagesAsync` exports selected PDF pages as images, `AssemblePdfAsync` combines supported PDFs, images, documents, folders, and ZIP archives, and `PdfPrintPlanner.Create` produces deterministic print-sheet placement plans.

Every workflow request runs with explicit input and output limits, cancellation, staged output validation, and a caller-selected collision policy. Passwords remain request-only values and are not copied into diagnostics or results. PDF comparison accepts a separate `ComparisonPdfPassword` when the two inputs use different credentials.

`PdfPrintRenderer.Prepare(document, request)` turns an authenticated `PdfDocument` snapshot and its `PdfPrintPlanRequest` into immutable PNG sheets. Display `prepared.Sheets[i].GetPng()` for review, then pass that same `PdfPreparedPrintDocument` to `IPdfPrinterService.SubmitAsync` with `PdfPrintDeliveryOptions`. Delivery does not reopen the source. `PdfPrinterService.GetPrintersAsync` lists installed queues; Windows uses GDI and macOS/Linux use the installed CUPS `lpstat`, `lpoptions`, and `lp` tools. Copies are collated, and duplex defaults to the printer's setting.

`GetPaperSourcesAsync(printerName)` lists the queue's reported paper sources. Set `PdfPrintDeliveryOptions.PaperSourceId` to one of those identifiers, or leave it null to retain the printer default. Identifiers belong to the queried queue; delivery rechecks an explicit selection before submitting. Windows reads driver bins and checks that the driver accepts the selected bin. CUPS discovers `InputSlot` or `media-source` choices exposed by `lpoptions -l`. Queues that expose no choices retain their default source.

Print preparation enforces PDF printing permissions, page and raster limits, cancellation, and a retained-output byte budget. Rendering above 150 DPI also requires high-quality print permission when restrictions apply. The output reflects the managed renderer's diagnostics and printer margins. `PdfPrintSubmission` is a queue acceptance receipt, not confirmation of physical delivery; `PdfPrintDeliveryException` means submission began and retrying could duplicate pages. Windows file-printer paths must be new local paths and are checked before submission; the driver controls the eventual write.

Applications that keep documents open can set `PublicationGuard` on `OfficeWorkflowRequest`, `PdfAssemblyRequest`, and `PdfPageImageExportRequest`. Implement `IOfficeWorkflowPublicationGuard.CanPublishAsync` to check live ownership of the supplied absolute destination. For directory outputs, check whether publication would replace a directory containing an owned document. The runner calls the guard after validating the staged artifact and checks every numbered candidate: a denied destination fails `Fail` or `Replace`, while `Rename` tries the next name. Cancellation and guard errors prevent publication. Calls can originate on worker threads, so UI hosts must dispatch ownership inspection to their UI thread. This is an application ownership check at publication time; it does not lock paths against concurrent external filesystem changes.

## Save protected or unencrypted PDF copies

```csharp
OfficeWorkflowResult protectedCopy = await OfficeWorkflow.ProtectPdf("report.pdf",
    new OfficeIMO.Pdf.PdfStandardEncryptionOptions(documentPassword) {
        OwnerPassword = ownerPassword,
        AllowedPermissions = OfficeIMO.Pdf.PdfStandardPermissions.Print
    })
    .To("protected.pdf")
    .RunAsync(cancellationToken: cancellationToken);

OfficeWorkflowResult unencryptedCopy = await OfficeWorkflow
    .RemovePdfProtection("protected.pdf", ownerPassword)
    .To("unencrypted.pdf")
    .RunAsync(cancellationToken: cancellationToken);
```

Typed requests use `ProtectPdf` with `OutputEncryption`, or `RemovePdfProtection`, and `PdfOwnerPassword` for existing protection. The owner password takes precedence over `PdfPassword` for these operations. To replace existing protection, use `ProtectPdf(...).WithPdfOwnerPassword(currentOwnerPassword)`. Settings are copied before execution; passwords are not included in results or reports.

These operations create separate copies, preserve their sources, and use the PDF engine's authorization and rewrite-preservation policy. Existing signatures or other protected document structures may prevent a rewrite. AES-256 is the default; AES-128 and explicitly selected legacy RC4 follow the canonical encryption options. Output verification checks the document-open password, page count, encryption state, permissions, metadata protection, and preservation report before publication. A null PDF metadata-encryption flag means the standard default of encrypted metadata.

Only the `Faithful` profile applies. Input snapshots, byte limits during generation, output conflicts, provider confirmation contracts, recovery, and final publication guards follow the same runner behavior as other single-output operations. Cancellation is forwarded through security preflight, graph rewriting, preservation inspection, and publication.

## Extract selected PDF pages

```csharp
OfficeWorkflowResult result = await OfficeWorkflow.ExtractPages("report.pdf", 5, 1, 2, 5)
    .To("selected-pages.pdf")
    .OnConflict(OfficeWorkflowConflictPolicy.Rename)
    .RunAsync(cancellationToken: cancellationToken);
```

The equivalent typed request uses `Operation = OfficeWorkflowOperation.ExtractPages` and `PageNumbers = [5, 1, 2, 5]`. Page numbers are one-based; order and intentional repeats are preserved, up to 100,000 selected pages. Extraction uses the PDF engine's page-preservation policy and supports only the `Faithful` profile. It creates a separate PDF and does not permit replacing the source.

The runner snapshots local and provider inputs, checks for source changes before publication, bounds output serialization, and reopens the generated PDF before publishing. `InputStream`, `OutputStream`, `PublicationGuard`, and the result's publication and recovery states follow the same contracts as other single-output workflows. Cancellation is observed before and after synchronous page extraction and during serialization; it cannot interrupt the PDF engine while that synchronous step is running.

## Save a certificate-signed PDF copy

Supply a caller-owned `IPdfExternalSigner` and `IPdfSignatureCryptographyProvider`. The PDF engine owns signature creation and inspection; the workflow captures the settings and publishes a separate output only after checking page count, signature structure, signature math, and document digests.

```csharp
OfficeWorkflowResult signedCopy = await OfficeWorkflow.SignPdf("report.pdf", signer,
    new OfficeIMO.Pdf.PdfExternalSignatureOptions {
        FieldName = "Approval",
        Reason = "Reviewed",
        VisibleAppearance = new() { PageNumber = 1, X = 36, Y = 36, Width = 180, Height = 48 }
    }, verifier)
    .To("signed-report.pdf")
    .RunAsync(cancellationToken: cancellationToken);
```

Keep the signer and verifier alive until the task completes. A host using `OfficeIMO.Security` can provide `PdfCmsExternalSigner` and `PdfCmsSignatureCryptographyProvider` adapters. The engine enforces the source document's permissions and certification policy; its current signing plan rejects documents that already contain a signature. Rejected requests leave the source and existing output unchanged.

`SignatureReport` reports certificate-chain, revocation, and timestamp evidence separately. Successful publication does not by itself establish certificate trust. A visible appearance identifies the signature on the page; it is not a substitute for cryptographic verification. On unconfirmed provider publication, the report describes the retained prepared artifact, not the destination's contents. Signing callbacks and cryptographic validation may be synchronous; cancellation is observed around those calls and prevents later publication.

## Split a PDF into consecutive parts

```csharp
PdfSplitWorkflowResult result = await runner.SplitPdfAsync(new PdfSplitWorkflowRequest {
    InputPath = "report.pdf",
    OutputDirectory = "report-parts",
    PagesPerDocument = 10,
    ConflictPolicy = OfficeWorkflowConflictPolicy.Rename
}, cancellationToken: cancellationToken);

foreach (PdfSplitFile file in result.Files) {
    Console.WriteLine($"{file.Path}: {file.PageCount} pages starting at source page {file.FirstSourcePage}");
}
```

The runner produces `part-001.pdf`, `part-002.pdf`, and subsequent parts in source order. Hosts can call `PdfSplitPlan.Create(pageCount, pagesPerDocument)` to preview the same filenames and ranges that execution uses. To split at chosen pages instead, such as top-level bookmarks, build `PdfSplitPlan.FromStarts(pageCount, starts)` from `PdfSplitStart(firstPage, title)` values and assign it to `PdfSplitWorkflowRequest.Plan`; parts are named `001-Title.pdf`, pages before the first start form a leading `part-001.pdf`, and the runner validates that the parts cover every source page exactly once in order and have unique file names. The runner generates and reopens one part at a time, checks the aggregate output budget before continuing, and publishes a local folder as a unit. `MaximumParts` limits the output count. Cancellation is checked between parts and during file operations; the PDF engine's synchronous generation of one part must finish before cancellation can stop it.

For provider folders, supply `DirectoryOutput` and explicitly choose `Replace`. Each part is written and verified individually. Inspect `Status`, `Files`, and `OutputRecoveries`: verified parts remain available if a later write fails. Local directory recovery locations appear in diagnostic details when an interrupted replacement needs attention. `InputStream` and `PublicationGuard` use the same source verification and live ownership contracts as other workflows.

## Read provider-backed inputs

Set `InputStream` on an `OfficeWorkflowRequest` when a file picker or storage provider supplies stream access. Keep `InputPath` as the original location or absolute URI, and supply the display filename for format routing:

```csharp
var request = new OfficeWorkflowRequest {
    Operation = OfficeWorkflowOperation.Convert,
    ConversionRouteId = "docx-pdf",
    InputPath = selectedLocation,
    InputStream = new OfficeWorkflowStreamInput(selectedName, openSelectedReadStream),
    OutputPath = outputPdfPath
};
OfficeWorkflowResult result = await runner.RunAsync(request, cancellationToken: cancellationToken);
```

`openSelectedReadStream` is a `Func<CancellationToken, Task<Stream>>`. It must return a fresh readable stream with the provider's permission scope each time. The runner closes every returned stream, stages a bounded private input for the document engine, and verifies the provider's SHA-256 again after host authorization and before publication. An optional `expectedSha256` constructor argument binds execution to contents captured when the user selected the input. Revoked access, changed contents, cancellation, and exceeded limits prevent publication. This is a point-in-time content check; providers do not offer a shared filesystem lock or atomic compare-and-replace contract.

For provider selections with a local path, the runner captures file identity while the read stream's access scope is active. It reopens that access for source/output checks and host authorization, then rejects physical source replacement even when the new file has identical contents. Local provider destinations also receive a read-scope check before writing; a new file may report `FileNotFoundException`, while other access failures prevent publication.

Comparison accepts `ComparisonStream`. Assembly accepts `SourceStreams`, keyed by the exact original entries in `Sources`, and preserves input order and display names. Its provider staging shares the total input byte budget. A provider HTML stream can use embedded resources; selecting it alone does not grant access to neighboring images or stylesheets. A selected ZIP can carry relative resources through the existing bounded archive intake.

For a selected provider folder, set `PdfAssemblyRequest.SourceDirectories` with an `OfficeWorkflowDirectoryInput` keyed by its original `Sources` entry. Its enumeration factory returns `OfficeWorkflowDirectoryEntry` values with a relative path, original location, and reopenable file input; a null input denotes a directory. Enumerate parents before children, obey the supplied recursion and traversal limits, and never follow links. Keep item references available until the runner returns. The runner preserves the relative tree for HTML resources, enforces aggregate entry and byte limits, rejects unsafe or colliding portable names, and rechecks both membership and file contents before publication. Each re-enumeration must reflect current provider state, including newly returned file objects at an existing location.

Page-image export accepts `PdfPageImageExportRequest.InputStream` and applies the same bounded staging and provider-content check before publishing its filesystem output folder. For print preview, use `PdfPrintPlanner.Create(document, request)` with an already opened `PdfDocument` to plan and render from the same snapshot. That overload uses the document's existing authentication and printing permissions; it does not reopen the request's input location.

Provider operations require an explicit output destination when they produce a file. Input staging is removed before publication or on failure; cleanup failures are reported. Report-only inspection and comparison may omit a destination.

## Write provider folders

Set `DirectoryOutput` on `PdfPageImageExportRequest` to write into a selected provider folder. Its `OfficeWorkflowDirectoryOutput` resolver receives each image filename and returns an `OfficeWorkflowDirectoryOutputFile` without modifying the provider. Existing children use their actual location. For new children whose location is assigned during creation, use the selected parent as the initial location and supply `OfficeWorkflowStreamOutput.PrepareDestination`; this callback creates the child after recovery is durable and returns its actual location. The runner authorizes that location before opening the write stream. Read factories must reopen the current child, not a cached copy of its bytes.

Provider folders require `Replace` and explicit consent. Each file is verified separately; cancellation or failure stops further writes without rolling back earlier ones. `Files` and `OutputBytes` describe verified outputs even when the batch does not complete. `OutputRecoveries` contains every retained recovery copy. Use these fields with `Status` rather than treating the folder as an atomic result. A recovery record for a newly created child identifies the selected parent and requested filename; successfully published files report their actual provider locations.

## Write provider-backed outputs

Set `OutputStream` on an `OfficeWorkflowRequest` or `PdfAssemblyRequest` to publish through a selected provider. Supply an `OfficeWorkflowStreamOutput` with the display filename, fresh read and write stream factories, and an `OfficeWorkflowOutputRecoveryStore` rooted in a private local directory. Set `OutputPath` to the original provider reference and `ConflictPolicy` to `Replace`. The host must obtain explicit consent for a direct write and for the required local recovery copy.

The runner validates the complete artifact and retains a verified local copy before preparing a new provider child or opening the write stream. It closes the write stream and reads the destination back to verify its SHA-256. A verified write returns `Completed` with the provider reference in `OutputPath`. A failure after the write starts returns `Unconfirmed`, leaves `OutputPath` unset, and exposes the retained copy through `Recovery`. Cancellation after the write starts also returns `Unconfirmed`; it does not prove that the destination is unchanged. Do not automatically retry these results.

The store defaults to a 1 GiB aggregate admission limit and at most 100 records. `GetRecoveries()` restores available records after restart, `VerifyAsync()` checks a copy before use, and `Discard()` removes a copy after explicit user action. Active publications are excluded from discovery. Successful or safely rejected writes remove their copies; cleanup failures are reported and can leave a recovery record. Before admitting another output, the store removes recognized incomplete records left before metadata publication, while preserving active leases and unfamiliar contents. Retained recovery copies do not expire automatically. Keep them outside normal output locations and require Save As when opening them for editing. Provider writes cannot guarantee atomic replacement, rollback, or exclusion of concurrent writers.

## Review and apply PDF redactions

Searchable PDF generation uses `OfficeWorkflowRunner.MakePdfSearchableAsync` with a `PdfSearchableWorkflowRequest` and a caller-owned `IOcrEngine`:

```csharp
var result = await new OfficeWorkflowRunner().MakePdfSearchableAsync(new() {
    InputPath = "scan.pdf",
    OutputPath = "searchable.pdf",
    ConflictPolicy = OfficeWorkflowConflictPolicy.Fail,
    Ocr = new OfficeIMO.Pdf.Ocr.PdfOcrMergeOptions { Language = "en", Dpi = 150 }
}, engine, cancellationToken);
```

The runner captures a bounded input snapshot, adds searchable text through `OfficeIMO.Pdf.Ocr`, and reopens the staged PDF before publication. It verifies source contents and local physical identity after recognition, then applies `PublicationGuard` and the selected conflict policy. The request also accepts `InputStream` and `OutputStream` with the same provider consent and recovery requirements described above. Inspect `Status`, `OutputPath`, and `Recovery` before opening or retrying an output. The engine remains owned by the caller.

Set `ReviewAsync` to pause before creating the text layer. The callback receives a `PdfSearchableOcrReview` and returns eligible word instances selected from that review. The shared PDF owner rejects foreign, duplicate, and policy-rejected selections. The destination remains untouched while review is pending, cancellation prevents publication, and source identity is checked again after the decision. Without a callback, the runner uses all eligible words.

Use `RecognizeImageAsync` for standalone images. It reads the image through `OfficeIMO.Reader.Image`, executes the selected engine through `OfficeIMO.Reader.Ocr`, and saves UTF-8 text. Its optional review callback receives the original image, recognition evidence, and recognized text, and returns the text to save:

```csharp
var result = await runner.RecognizeImageAsync(new ImageOcrWorkflowRequest {
    InputPath = "invoice.png",
    OutputPath = "invoice.txt",
    Ocr = new OfficeIMO.Reader.OfficeDocumentOcrExecutionOptions { Language = "eng" },
    ReviewAsync = (review, token) => Task.FromResult(review.Text)
}, engine, cancellationToken);
```

The callback can present a preview and accept corrections. Until it returns, the destination is untouched. Empty recognition remains visible in diagnostics; failed or skipped recognition does not publish a partial text file. Corrected text is bounded by `Limits.MaximumOutputBytes`, reopened before publication, and protected by the same source identity, conflict, provider consent, and recovery contracts as PDF output.

`RunOcrSessionAsync` accepts an ordered collection of `OfficeOcrSessionRequest` items containing either request type. It snapshots request settings, uses one caller-owned engine sequentially, and protects every selected source from every output. Each item has a unique caller id and a distinct output destination. Progress and terminal result callbacks let a host show completed outputs while later items await review. Cancellation retains completed outputs and returns cancelled outcomes for unstarted items. A retry should contain only the explicitly selected failed or cancelled items; an `Unconfirmed` result stops the remaining items and requires checking the destination and recovery copy first. Pass previously completed output locations through `protectedOutputPaths` when retrying a subset, including when a provider resolves a different destination during publication. Also pass every retained session source as an `OfficeWorkflowProtectedSource` through `protectedInputs`, including its provider stream access. This preserves original inputs of completed items while retrying or adding work.

Redaction uses a separate versioned plan/review/apply contract. Planning produces privacy-safe candidate identifiers and geometry. Application re-plans the exact source and recipe, requires every current candidate to be explicitly approved or rejected, applies only approved candidates, and publishes only after native and configured OCR verification succeeds.

```csharp
var recipe = new PdfRedactionRecipe();
recipe.Rules.Add(new PdfRedactionRule {
    Name = "account-number",
    Kind = PdfRedactionRuleKind.Literal,
    Value = "Account: 123-45-6789",
    ContentScope = PdfRedactionContentScope.TextAndUnderlay,
    AppearanceMode = PdfRedactionAppearanceMode.QuantizedWidth
});

var runner = new OfficeWorkflowRunner();
PdfRedactionWorkflowResult plan = await runner.RunRedactionAsync(
    new PdfRedactionWorkflowRequest {
        Mode = PdfRedactionWorkflowMode.PlanOnly,
        InputPath = "contract.pdf",
        EvidencePath = "contract.plan.json",
        Recipe = recipe
    });

var decisions = new PdfRedactionDecisionManifest {
    SourceSha256 = plan.SourceSha256,
    RecipeSha256 = plan.RecipeSha256,
    ApprovedCandidateIds = plan.Candidates.Select(candidate => candidate.Id).ToList()
};

PdfRedactionWorkflowResult applied = await runner.RunRedactionAsync(
    new PdfRedactionWorkflowRequest {
        Mode = PdfRedactionWorkflowMode.ApplyAndVerify,
        InputPath = "contract.pdf",
        OutputPath = "contract-redacted.pdf",
        EvidencePath = "contract-redacted.evidence.json",
        Recipe = recipe,
        Decisions = decisions
    });
```

Rule and explicit-region names are stable, non-sensitive evidence identifiers. `ContentScope` decides whether a reviewed area removes only text or also intersecting underlay content. `AppearanceMode` independently controls the privacy of the visible mark: exact, nearby-merged, quantized-width, or full-line. Recipe, decision, and batch JSON reject unknown members so misspelled policy fields cannot silently fall back to defaults.

The schemas are `officeimo.pdf.redaction.recipe.v1`, `officeimo.pdf.redaction.plan.v1`, `officeimo.pdf.redaction.decisions.v1`, `officeimo.pdf.redaction.result.v1`, `officeimo.pdf.redaction.batch-request.v1`, and `officeimo.pdf.redaction.batch.v1`. Persisted `PdfRedactionWorkflowRecord` JSON omits matched text, extracted text, passwords, OCR payloads, provider options, host paths, and caller request identifiers. The in-memory operational result still carries paths and request correlation for host UX. Evidence retains rule names, policies, hashes, counts, complete atomic candidate geometry, stable issue codes, one-way SHA-256 fingerprints of provider/model/language values, and OCR confidence. Raw provider-returned metadata is never persisted, so document text or credentials cannot become evidence metadata even when they contain only identifier characters. Encrypted input requires an explicit reject, decrypt, or decrypt-and-reencrypt policy with runtime-only owner credentials. Zero-area verification of a re-encrypted output also requires the trusted output SHA-256 from prior apply evidence.

Signed input uses an explicit `SignaturePolicy`. The default rejects it. `CreateUnsignedDerivative` removes invalidated signature structures through a full rewrite before planning and records source/output signature counts; `CreateAndSignDerivative` additionally requires a runtime `IPdfExternalSigner` and can cryptographically validate the new signature through an optional `IPdfSignatureCryptographyProvider`. The output is always a separate artifact. Runtime `ExternalValidators` accept `IPdfRedactionCancellationAwareExternalValidator` implementations that bind independent parser, renderer, or forensic checks to the final bytes; their names are retained in evidence, cancellation stops the workflow before publication, and any rejection prevents publication.

Single-item evidence, per-output bytes, batch items, concurrency, and aggregate prepared output/evidence bytes have independent limits. Batch preparation reserves each in-flight item's configured worst-case size and fails before publication when the aggregate ceiling cannot be honored; successful items are reclassified as unpublished if any sibling fails.

The file-set overload deterministically selects PDFs and mirrors their relative directories into separate output, evidence, and decision roots:

```csharp
PdfRedactionBatchResult batch = await runner.RunRedactionBatchAsync(
    new PdfRedactionBatchRequest {
        Mode = PdfRedactionWorkflowMode.PlanOnly,
        InputRoot = "incoming",
        EvidenceRoot = "review-evidence",
        ManifestPath = "review-evidence/batch.json",
        Recipe = recipe,
        PublicationPolicy = PdfRedactionBatchPublicationPolicy.AtomicAll
    });
```

`RunRedactionBatchAsync` prepares every bounded item before atomic publication with configurable concurrency, stages every file beside its destination, and rolls back already published files if an ordinary publication failure occurs. `ContinuePerItem` instead publishes successful items independently and records failures in the consolidated manifest. Batch destinations must be portable-case unique, remain physically outside the input root, and use one fail-or-replace conflict policy. Recursive discovery does not follow reparse points, and explicit linked inputs must still resolve beneath the physical input root. This is an in-process publication transaction, not a filesystem-wide crash transaction.

## Inspect and remove provenance

The provenance workflow keeps format logic in its owning package. `OfficeIMO.Word`, `OfficeIMO.Excel`, `OfficeIMO.PowerPoint`, `OfficeIMO.Visio`, `OfficeIMO.OpenDocument`, `OfficeIMO.Epub`, `OfficeIMO.Pdf`, `OfficeIMO.Html`, and `OfficeIMO.Markdown` handle their formats; `OfficeIMO.Core` handles supported images and structured text. Consumers can discover the exact extension, structural format, owner, operation, memory-only, and browser contract through `OfficeProvenanceWorkflowCatalog.All`, `ToJson()`, or `ToMarkdown()`.

The workflow requires a registered extension and matching structural format. It does not infer ownership for unknown extensions or generic containers. Applications that already own such a format context can call the lower-level `OfficeProvenanceInspector` API directly for signature-based inspection.

```csharp
using OfficeIMO.Workflows;

var runner = new OfficeWorkflowRunner();

OfficeProvenanceWorkflowResult inspection = await runner.RunProvenanceAsync(
    new OfficeProvenanceWorkflowRequest {
        Operation = OfficeProvenanceWorkflowOperation.Inspect,
        InputPath = "report.docx"
    });

OfficeProvenanceWorkflowResult removal = await runner.RunProvenanceAsync(
    new OfficeProvenanceWorkflowRequest {
        Operation = OfficeProvenanceWorkflowOperation.Remove,
        InputPath = "report.docx",
        OutputPath = "report.cleaned.docx",
        ExpectedInputSha256 = inspection.InputSha256,
        ConflictPolicy = OfficeWorkflowConflictPolicy.Fail
    });
```

`Assess` combines the owner-specific structural report with exact Unicode findings and optional `IOfficeProvenanceVerifier` / `IOfficeProvenanceSignalDetector` services supplied to the runner. It preserves each provider's result and does not infer a universal authorship verdict.

Removal is strict by default. It removes only selected, structurally valid carriers and blocks a package-signature-invalidating save unless the caller explicitly selects `OfficeSignatureMutationPolicy.RemoveInvalidatedSignatures`. The output is written to a sibling staging file, reopened through the same format owner, checked against the removal report, and only then published under the requested conflict policy. Generic ZIP packages and renamed package subtypes are rejected because the workflow has no matching registered format owner for them.

`OfficeProvenanceReportSerializer.Serialize(result)` produces the same `officeimo.provenance.result.v2` document used by the CLI, Studio report export, and browser provenance download. `SerializeBatch(results)` uses `officeimo.provenance.batch.v2`. Reports retain structured evidence and diagnostics, string enum values, coverage notes, input/output SHA-256 hashes, and explicit check states. Assessment reports distinguish disabled or unsupported Unicode inspection from a completed empty report and distinguish an absent provider from verification that ran.

Pass `ExpectedInputSha256` from a reviewed result when a later action must use the same source bytes. The runner compares the immutable input snapshot before any mutation. `PublicationGuard` applies the host's live ownership check to each final destination; Fail/Replace reject an owned path and Rename skips it.

`OfficeProvenanceAudit.RunAsync(new OfficeProvenanceAuditRequest { Inputs = ["documents"], Include = ["*.html"], MaximumItems = 1000 })` discovers and assesses a bounded set without modifying it. Discovery is recursive by default, excludes symbolic links and common generated/VCS directories, and fails on an empty selection or exceeded bounds. Explicit files retain ordinary workflow errors. `OfficeProvenanceAudit.HasFindings(result, carriers: false, dangerousText: true)` evaluates an evidence policy; callers must handle execution failures separately. `OfficeProvenanceSarif.Serialize(results)` exports the same evidence and failures as SARIF 2.1.0.

Use `RunProvenanceBatchAsync` for bounded sequential batches. Sequential execution keeps parser and provider resource use predictable, while per-request progress includes an overall batch fraction.


For a memory-only host, `OfficeProvenanceBufferWorkflow.Inspect(bytes, fileName, options)` and `Remove(bytes, fileName, removalOptions)` use the same catalog and format owners without opening paths or following remote references. Read the qualified extensions from `OfficeProvenanceWorkflowCatalog.MemoryOnlyExtensions`; the current families are JPEG, PNG, WebP, PDF, DOCX, XLSX, and PPTX. Removal returns a separate result and re-inspects its bytes before returning. Specify limits appropriate to the host; a browser should use tighter limits than a local batch runner.

```csharp
var inspection = OfficeProvenanceBufferWorkflow.Inspect(inputBytes, "report.docx");
var result = OfficeProvenanceBufferWorkflow.Remove(inputBytes, "report.docx");
byte[] cleanedCopy = result.ToArray();
// Inspect result.After and result.Changes before presenting the copy to the user.
```

For memory-only report export, pass the inspected bytes and report to `OfficeProvenanceReportSerializer.FromBuffer(fileName, bytes, inspection, removal)` and serialize the returned result. These factories do not read paths or verify cryptographic authenticity.

`OfficeTextIntegrityReview` in Core owns source-bound text selections and encoding-preserving export. `OfficeTextIntegrityReportSerializer.Serialize(review, review.Text, fileName, selectedIndices)` exports exact findings, selected occurrence indices, UTF-16 offset units, source hashes, encoding/BOM information, and the selected-copy digest. It does not include the full source text.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Create | 1 | 0 | 0 | 0 | 0 | 0 |
| Read | 1 | 0 | 0 | 0 | 0 | 0 |
| Edit | 1 | 0 | 0 | 0 | 0 | 0 |
| Preserve | 1 | 0 | 0 | 0 | 0 | 0 |
| Validate | 0 | 1 | 0 | 0 | 0 | 0 |
| Convert | 0 | 2 | 0 | 0 | 0 | 0 |
| Export | 1 | 0 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.Workflows` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
