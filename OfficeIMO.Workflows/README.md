# OfficeIMO.Workflows

`OfficeIMO.Workflows` is the reusable local orchestration layer for OfficeIMO document jobs. It composes the existing first-party conversion, PDF, and provenance APIs behind typed requests, bounded execution, cooperative cancellation, collision policies, atomic output publication, and post-write validation.

The package does not add a second document or PDF engine. Desktop applications, command-line tools, and services can share this workflow contract while keeping their user-interface and hosting code thin.

## Reference from source

When working from an OfficeIMO source checkout, reference the workflow project directly:

```xml
<ProjectReference Include="..\OfficeIMO.Workflows\OfficeIMO.Workflows.csproj" />
```

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

## Convert a document

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

`PdfPrintRenderer.Prepare(document, request)` turns an authenticated `PdfDocument` snapshot and its `PdfPrintPlanRequest` into immutable PNG sheets. Display `prepared.Sheets[i].GetPng()` for review, then pass that same `PdfPreparedPrintDocument` to `IPdfPrinterService.SubmitAsync` with `PdfPrintDeliveryOptions`. Delivery does not reopen the source. `PdfPrinterService.GetPrintersAsync` lists installed queues; Windows uses GDI and macOS/Linux use the installed CUPS `lpstat` and `lp` tools. Copies are collated, and duplex defaults to the printer's setting.

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

The runner produces `part-001.pdf`, `part-002.pdf`, and subsequent parts in source order. Hosts can call `PdfSplitPlan.Create(pageCount, pagesPerDocument)` to preview the same filenames and ranges that execution uses. The runner generates and reopens one part at a time, checks the aggregate output budget before continuing, and publishes a local folder as a unit. `MaximumParts` limits the output count. Cancellation is checked between parts and during file operations; the PDF engine's synchronous generation of one part must finish before cancellation can stop it.

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

The provenance workflow keeps format logic in its owning package. `OfficeIMO.Word`, `OfficeIMO.Excel`, `OfficeIMO.PowerPoint`, `OfficeIMO.Visio`, `OfficeIMO.OpenDocument`, `OfficeIMO.Epub`, `OfficeIMO.Pdf`, `OfficeIMO.Html`, and `OfficeIMO.Markdown` handle their formats; `OfficeIMO.Core` handles supported images and structured text. Consumers can discover the exact extension-to-owner map through `OfficeProvenanceWorkflowCatalog.All`.

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
        ConflictPolicy = OfficeWorkflowConflictPolicy.Fail
    });
```

`Assess` combines the owner-specific structural report with exact Unicode findings and optional `IOfficeProvenanceVerifier` / `IOfficeProvenanceSignalDetector` services supplied to the runner. It preserves each provider's result and does not infer a universal authorship verdict.

Removal is strict by default. It removes only selected, structurally valid carriers and blocks a package-signature-invalidating save unless the caller explicitly selects `OfficeSignatureMutationPolicy.RemoveInvalidatedSignatures`. The output is written to a sibling staging file, reopened through the same format owner, checked against the removal report, and only then published under the requested conflict policy. Generic ZIP packages can be inspected but are not mutated without a registered format owner.

Use `RunProvenanceBatchAsync` for bounded sequential batches. Sequential execution keeps parser and provider resource use predictable, while per-request progress includes an overall batch fraction.
