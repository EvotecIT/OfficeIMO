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

`IWorkSourceDocument.Open` reads and bounds the source independently of destination policy. `ToPowerPointPresentation` returns a complete editable presentation and rejects partial reconstruction or preview output; `ToPowerPointPresentationResult` also exposes the typed Keynote projection, diagnostics, preserved source records, and exact editable-versus-visual-fallback result. `PowerPointIWorkConverter.ConvertKeynoteToPowerPoint*` provides equivalent path and stream convenience entry points.

Qualified opaque slide background colors and explicit no-fill overrides survive PPTX save/reopen, including selected style inheritance. Unsupported backgrounds require visual fallback or explicit partial conversion and retain source diagnostics. See the [background contract and native evidence](../Docs/officeimo.iwork-support-matrix.md#keynote-slide-backgrounds) for the supported color subset and remaining master/export limits.

Table-region defaults and selected text styles preserve supported fonts, emphasis, colors and paragraph alignment in PPTX, including empty cells. Explicit rich-text formatting takes precedence. Table paragraph pagination flags require the partial policy and produce `IWORK_KEYNOTE_PARAGRAPH_PAGINATION_OMITTED`; strict conversion uses visual fallback.

Selected native cell padding becomes PowerPoint table-cell margins, and top/middle/bottom alignment becomes the cell anchor. Values must fit the PPTX margin range; the partial policy permits EMU rounding with a precision diagnostic.

Supported selected and unbanded region fills override the PPTX table theme, including explicit no-fill and unstored empty cells. Selected fills take precedence over region defaults. Unsupported fills and banding remain diagnosed on the source projection; banding and complete native appearance remain unqualified.

Individual table row heights and column widths are carried into PPTX. If their total differs from the table’s drawable extent, the adapter scales them proportionally and reports `IWORK_KEYNOTE_TABLE_SIZING_SCALED` as an approximation. Measurements are quantized to EMUs with a separate precision diagnostic.

Supported numeric table formats become editable PPTX text: decimal precision, grouping, percentages, scientific notation, fractions, currency-code prefixes and negative-value parentheses. Red negative formats become run color. The report identifies display approximation because locale-specific symbols, automatic precision and source appearance can differ. Qualified date/time patterns and fixed day-only or hour/minute durations also become editable table text through the shared numeric, calendar and elapsed formatters; see the [five-pattern source contract](../Docs/officeimo.iwork-support-matrix.md#conversion-acceptance-and-fidelity). Raw values and formula caches remain on the source projection; formatting does not change its raw display properties. Native Keynote export qualification remains outside this bounded contract.

Qualified source-hidden table rows and columns remain visible in editable PPTX output. Conversion requires `AllowPartialEditableReconstruction = true` and reports `IWORK_KEYNOTE_TABLE_VISIBILITY_OMITTED`; strict conversion uses visual fallback when available. Hidden positions and cell content remain on the typed source projection.

Choose the acceptance policy explicitly when source details cannot be represented:

```csharp
var options = new IWorkConversionOptions {
    Mode = IWorkConversionMode.Auto,
    AllowPartialEditableReconstruction = true,
    RequireCompleteVisualCoverage = true
};
```

This retains bounded recoverable editable content and reports incomplete details. If editable output cannot be produced, it rejects a first-page or composite preview. `AllowPartialEditableReconstruction` defaults to `false`; `RequireCompleteVisualCoverage` defaults to `true`. To accept an incomplete preview, use the result API with `RequireCompleteVisualCoverage = false` and inspect its coverage and fidelity report. Value-only APIs reject partial and preview output even when these options permit it. `result.RequireCompleteEditableReconstruction()` returns the destination after checking assessed content completeness and disposes rejected output; it does not establish identical appearance or field-level fidelity. `Report.RequireNoLoss()` also rejects unassessed record fidelity. These policies do not bypass source limits or destination safety checks.

Under the partial policy, recovered slides remain editable when source paragraph pagination flags cannot be represented. The report retains the pagination diagnostic and the partial-reconstruction finding.

The path and stream convenience APIs accept cancellation after the options:

```csharp
using var cancellation = new CancellationTokenSource();
using KeynoteToPowerPointResult cancellable = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(
    "source.key", readOptions: null, conversionOptions: options,
    cancellationToken: cancellation.Token);
```

This token governs loading, projection, and destination construction. It also governs later projections from `cancellable.Source`; use `cancellable.Source.WithCancellation(newToken)` to reuse the loaded package with a new operation token. This replaces the previous token while sharing source bytes and parsed messages. Saving is a separate destination-owner operation.

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
