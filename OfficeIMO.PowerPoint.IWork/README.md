# OfficeIMO.PowerPoint.IWork

`OfficeIMO.PowerPoint.IWork` is the opt-in adapter for importing modern Apple Keynote files into editable `OfficeIMO.PowerPoint` presentations. Installing `OfficeIMO.PowerPoint` alone does not add the iWork reader.

```bash
dotnet add package OfficeIMO.PowerPoint.IWork
```

```csharp
using OfficeIMO.IWork;
using OfficeIMO.PowerPoint.IWork;

IWorkSourceDocument source = IWorkSourceDocument.Open("source.key");
using KeynoteToPowerPointResult result = source.ToPowerPointPresentationResult(
    new IWorkConversionOptions { Mode = IWorkConversionMode.Auto });

Console.WriteLine(result.Report.ProjectionKind);
Console.WriteLine(result.HasLoss);
result.Value.Save("converted.pptx");
```

`IWorkSourceDocument.Open` reads and bounds the source independently of destination policy. `ToPowerPointPresentation` returns the converted presentation directly; `ToPowerPointPresentationResult` also exposes the typed Keynote projection, diagnostics, preserved source records, and exact editable-versus-visual-fallback result. `PowerPointIWorkConverter.ConvertKeynoteToPowerPoint*` provides equivalent path and stream convenience entry points.

Table-region defaults and selected text styles preserve supported fonts, emphasis, colors and paragraph alignment in PPTX, including empty cells. Explicit rich-text formatting takes precedence. Table paragraph pagination flags require the partial policy and produce `IWORK_KEYNOTE_PARAGRAPH_PAGINATION_OMITTED`; strict conversion uses visual fallback.

Selected native cell padding becomes PowerPoint table-cell margins, and top/middle/bottom alignment becomes the cell anchor. Values must fit the PPTX margin range; the partial policy permits EMU rounding with a precision diagnostic.

Supported selected and unbanded region fills override the PPTX table theme, including explicit no-fill and unstored empty cells. Selected fills take precedence over region defaults. Unsupported fills and banding remain diagnosed on the source projection; banding and complete native appearance remain unqualified.

Individual table row heights and column widths are carried into PPTX. If their total differs from the table’s drawable extent, the adapter scales them proportionally and reports `IWORK_KEYNOTE_TABLE_SIZING_SCALED` as an approximation. Measurements are quantized to EMUs with a separate precision diagnostic.

Choose the acceptance policy explicitly when source details cannot be represented:

```csharp
var options = new IWorkConversionOptions {
    Mode = IWorkConversionMode.Auto,
    AllowPartialEditableReconstruction = true,
    RequireCompleteVisualCoverage = true
};
```

This retains bounded recoverable editable content and reports incomplete details. If editable output cannot be produced, it rejects a first-page or composite preview. Both settings default to `false`. Inspect `Report.IsPartialEditableReconstruction` before accepting the output; `Report.RequireCompleteEditableReconstruction()` rejects explicitly partial reconstruction. `Report.RequireNoLoss()` also rejects unassessed record fidelity. These policies do not bypass source limits or destination safety checks.

Under the partial policy, recovered slides remain editable when source paragraph pagination flags cannot be represented. The report retains the pagination diagnostic and the partial-reconstruction finding.

The path and stream convenience APIs accept cancellation after the options:

```csharp
using var cancellation = new CancellationTokenSource();
using KeynoteToPowerPointResult cancellable = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(
    "source.key", readOptions: null, conversionOptions: options,
    cancellationToken: cancellation.Token);
```

This token governs loading, projection, and destination construction. It also governs later projections from `cancellable.Source`; reopen the source with a new token after cancellation. Saving is a separate destination-owner operation.

The adapter directly depends on `OfficeIMO.Core`, `OfficeIMO.IWork`, and `OfficeIMO.PowerPoint`. It does not add iWork support to the default PowerPoint package graph.

See the [iWork support matrix](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.iwork-support-matrix.md) for supported structures and conversion limits.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Convert | 0 | 1 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.PowerPoint.IWork` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
