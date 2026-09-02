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
