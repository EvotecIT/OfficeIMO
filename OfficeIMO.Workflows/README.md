# OfficeIMO.Workflows

`OfficeIMO.Workflows` is the reusable local orchestration layer for OfficeIMO document jobs. It composes the existing first-party conversion and PDF APIs behind typed requests, bounded execution, cooperative cancellation, collision policies, atomic output publication, and post-write validation.

The package does not add a second document or PDF engine. Desktop applications, command-line tools, and services can share this workflow contract while keeping their user-interface and hosting code thin.

## Reference from source

`OfficeIMO.Workflows` is new in this development line and does not yet have a published NuGet version. Until its first package release is available, reference the project from an OfficeIMO source checkout:

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
