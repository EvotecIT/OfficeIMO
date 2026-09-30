# OfficeIMO.Workflows.IWork

Opt-in Pages-to-Word, Numbers-to-Excel, and Keynote-to-PowerPoint conversion through the shared OfficeIMO workflow runner.

```csharp
using OfficeIMO.Workflows;
using OfficeIMO.Workflows.IWork;

OfficeWorkflowRunner runner = IWorkWorkflow.CreateRunner();
OfficeWorkflowResult result = await runner.RunAsync(new OfficeWorkflowRequest {
    InputPath = "budget.numbers",
    OutputPath = "budget.xlsx",
    Operation = OfficeWorkflowOperation.Convert,
    ConversionRouteId = "numbers-xlsx"
});
```

The runner captures the source, bounds input and output, reopens the staged destination through its format owner, verifies the source before publication, and applies the requested conflict policy. The `SourceSnapshot` diagnostic identifies the captured ZIP bytes by SHA-256, length, and selected filename. `ConversionEvidence` retains typed fidelity categories and compact producer, projection, and coverage facts. Successful publication does not establish complete source fidelity or visual equivalence.

The default acceptance policy rejects explicitly partial editable reconstruction and previews without known full-document coverage. Pass `IWorkConversionOptions` to `CreateRunner` to choose editable or visual output, permit partial reconstruction, or normalize Numbers worksheet names. These options are copied when routes are registered. Each request can instead supply its own `IWorkWorkflowSettings`; the runner snapshots them before opening the source, so later caller edits cannot change a pending conversion. Destination renames, approximations, omissions, and unassessed source records remain reported.

```csharp
var settings = new IWorkWorkflowSettings();
settings.ConversionOptions.Mode = OfficeIMO.IWork.IWorkConversionMode.VisualOnly;
// Explicitly accept a preview that may cover only part of the source.
settings.ConversionOptions.RequireCompleteVisualCoverage = false;
var request = new OfficeWorkflowRequest {
    InputPath = "report.pages",
    OutputPath = "report.docx",
    Operation = OfficeWorkflowOperation.Convert,
    ConversionRouteId = "pages-docx",
    RegisteredConversionSettings = settings
};
OfficeWorkflowResult preview = await runner.RunAsync(request);
```

Routes accept ZIP files and provider ZIP streams. Directory packages are supported by the source reader and Reader adapter; workflow conversion currently requires a ZIP input. See the [iWork support matrix](../Docs/officeimo.iwork-support-matrix.md) for format limits and the [workflow guide](../OfficeIMO.Workflows/README.md) for publication and provider contracts.
