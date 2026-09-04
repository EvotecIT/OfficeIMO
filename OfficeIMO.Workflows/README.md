# OfficeIMO.Workflows

`OfficeIMO.Workflows` is the reusable local orchestration layer for OfficeIMO document jobs. It composes the existing first-party conversion, PDF, and provenance APIs behind typed requests, bounded execution, cooperative cancellation, collision policies, atomic output publication, and post-write validation.

The package does not add a second document or PDF engine. Desktop applications, command-line tools, and services can share this workflow contract while keeping their user-interface and hosting code thin.

## Reference from source

When working from an OfficeIMO source checkout, reference the workflow project directly:

```xml
<ProjectReference Include="..\OfficeIMO.Workflows\OfficeIMO.Workflows.csproj" />
```

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

`RunAsync` also exposes PDF inspection, comparison, optimization, repair planning, repair, and sanitization through typed operations. `ExportPdfPagesAsync` exports selected PDF pages as images, `AssemblePdfAsync` combines supported PDFs, images, documents, folders, and ZIP archives, and `PdfPrintPlanner.Create` produces deterministic print-sheet placement plans.

Every request runs with explicit input and output limits, cancellation, staged output validation, and a caller-selected collision policy. Passwords remain request-only values and are not copied into diagnostics or results. PDF comparison accepts a separate `ComparisonPdfPassword` when the two inputs use different credentials.

## Review and apply PDF redactions

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

The schemas are `officeimo.pdf.redaction.recipe.v1`, `officeimo.pdf.redaction.plan.v1`, `officeimo.pdf.redaction.decisions.v1`, `officeimo.pdf.redaction.result.v1`, `officeimo.pdf.redaction.batch-request.v1`, and `officeimo.pdf.redaction.batch.v1`. Persisted `PdfRedactionWorkflowRecord` JSON omits matched text, extracted text, passwords, OCR payloads, provider options, host paths, and caller request identifiers. The in-memory operational result still carries paths and request correlation for host UX. Evidence retains rule names, policies, hashes, counts, complete atomic candidate geometry, stable issue codes, provider/model/language identifiers, and OCR confidence. Encrypted input requires an explicit reject, decrypt, or decrypt-and-reencrypt policy with runtime-only owner credentials. Zero-area verification of a re-encrypted output also requires the trusted output SHA-256 from prior apply evidence.

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

`RunRedactionBatchAsync` prepares every bounded item before atomic publication with configurable concurrency, stages every file beside its destination, and rolls back already published files if an ordinary publication failure occurs. `ContinuePerItem` instead publishes successful items independently and records failures in the consolidated manifest. Batch destinations must be portable-case unique, stay outside the input root, and use one fail-or-replace conflict policy. This is an in-process publication transaction, not a filesystem-wide crash transaction.

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
