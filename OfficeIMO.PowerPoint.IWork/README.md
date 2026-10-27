# OfficeIMO.PowerPoint.IWork

`OfficeIMO.PowerPoint.IWork` is the opt-in adapter for importing modern Apple Keynote files into editable `OfficeIMO.PowerPoint` presentations. Installing `OfficeIMO.PowerPoint` alone does not add the iWork reader.

```bash
dotnet add package OfficeIMO.PowerPoint.IWork
```

```csharp
using OfficeIMO.IWork;
using OfficeIMO.PowerPoint.IWork;

IWorkSourceDocument source = IWorkSourceDocument.Open("source.key");
using KeynoteToPowerPointResult result = source.ToPowerPointDocumentResult(
    new IWorkConversionOptions { Mode = IWorkConversionMode.Auto });

Console.WriteLine(result.Report.ProjectionKind);
Console.WriteLine(result.HasLoss);
result.Value.Save("converted.pptx");
```

`IWorkSourceDocument.Open` reads and bounds the source independently of destination policy. `ToPowerPointDocument` returns a complete editable presentation and rejects partial reconstruction or preview output; `ToPowerPointDocumentResult` also exposes the typed Keynote projection, diagnostics, preserved source records, and exact editable-versus-visual-fallback result. `PowerPointIWorkConverter.ConvertKeynoteToPowerPoint*` provides equivalent path and stream convenience entry points.

Qualified opaque slide background colors and explicit no-fill overrides survive PPTX save/reopen, including selected style inheritance. Unsupported backgrounds require visual fallback or explicit partial conversion and retain source diagnostics. See the [background contract and native evidence](../Docs/officeimo.iwork-support-matrix.md#keynote-slide-backgrounds) for the supported color subset and remaining master/export limits.

Selected text frames preserve recovered placeholder geometry, inherited margins, vertical alignment and shrink-to-fit. Qualified character lists preserve marker placement and relative size with a uniform known text font size. Interacting paragraph indents, mixed font sizes and marker scales outside whole percentages from 25% to 400% remain unsupported. The [strict everyday subset](../Docs/officeimo.iwork-support-matrix.md#qualified-everyday-subset) records native source/export and rendered evidence for a two-slide deck, including presenter notes and a wrapped fixed-frame probe.

Native `?slideid=<UUID>` targets resolve through selected slide-node identities in package metadata, and `?slide=next` resolves to the next source position. The adapter writes stable PPTX slide links while the source projection retains the original values. Missing or ambiguous identities and a missing next slide remain unsupported, including under the partial policy. Finite text-frame insets round to the PPTX EMU grid with `IWORK_KEYNOTE_PPTX_PRECISION` evidence; out-of-range values remain rejected. The [three-slide native reference](../OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/keynote-slide-links-v15.4.json) qualifies default editable conversion, save/reopen and full-slide PDF link targets for Keynote 15.4.

Table-region defaults and selected text styles preserve supported fonts, emphasis, colors and paragraph alignment in PPTX, including empty cells. Explicit rich-text formatting takes precedence. Table paragraph pagination flags require the partial policy and produce `IWORK_KEYNOTE_PARAGRAPH_PAGINATION_OMITTED`; strict conversion uses visual fallback.

Selected native cell padding becomes PowerPoint table-cell margins, and top/middle/bottom alignment becomes the cell anchor. Values must fit the PPTX margin range; the partial policy permits EMU rounding with a precision diagnostic.

Supported selected and unbanded region fills override the PPTX table theme, including explicit no-fill and unstored empty cells. Selected fills take precedence over region defaults. Unsupported fills and banding remain diagnosed on the source projection; banding and complete native appearance remain unqualified.

Individual table row heights and column widths are carried into PPTX. If their total differs from the table’s drawable extent, the adapter scales them proportionally and reports `IWORK_KEYNOTE_TABLE_SIZING_SCALED` as an approximation. Measurements are quantized to EMUs with a separate precision diagnostic.

Supported numeric table formats become editable PPTX text: decimal precision, grouping, percentages, scientific notation, fractions, currency-code prefixes and negative-value parentheses. Red negative formats become run color. The report identifies display approximation because locale-specific symbols, automatic precision and source appearance can differ. Qualified date/time patterns and fixed abbreviated day-only, hour/minute, hour/second, minute/second or second-only durations also become editable table text through the shared numeric, calendar and elapsed formatters; see the [five-pattern source contract](../Docs/officeimo.iwork-support-matrix.md#conversion-acceptance-and-fidelity). Raw values and formula caches remain on the source projection; formatting does not change its raw display properties. Native Keynote export qualification remains outside this bounded contract.

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

Qualified non-tiered decimal levels 0 and 1 retain starts, restarts and parent resets. Resuming after ordinary text, returning to a parent level or changing the qualified marker font/size restates the recovered counter for the receiving application. Counters outside 1–32767 are rejected before destination mutation. The bounded decimal marker profile stores a whole-point marker size independently of uniform body-run size through `BulletSizePoints`, preserving 11-point marker declarations with 12-point text without percentage rounding. Other marker profiles keep their existing measurement and fidelity checks; native Keynote rendering retains its own qualification boundary.

The [counter component references](../OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/list-counter-mapping-v15.4.json) retain every counter and paragraph after Keynote 15.4 imports the saved PPTX and exports PDF. They cover a child restart followed by a parent reset, continuation through ordinary text and a marker-font change. The synthetic partial-conversion cases use a resized table frame to expose all paragraphs. The explicit-restart case displays a missing-font warning; full layout and font appearance remain unqualified. Composed table PDFs retain the markers and report their inline layout approximation.

Keep-lines is inactive in qualified unlinked, horizontal, single-column fixed frames. Keep-with-next, page breaks, unknown frame layouts, linked flows, presenter notes and table-cell pagination retain the strict checks. Under the partial policy, recovered slides remain editable when those paragraph flags cannot be represented; the report retains the pagination diagnostic and partial-reconstruction finding.

The path and stream convenience APIs accept cancellation after the options:

```csharp
using var cancellation = new CancellationTokenSource();
using KeynoteToPowerPointResult cancellable = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(
    "source.key", readOptions: null, conversionOptions: options,
    cancellationToken: cancellation.Token);
```

This token governs loading, projection, and destination construction. It also governs later projections from `cancellable.Source`; use `cancellable.Source.WithCancellation(newToken)` to reuse the loaded package with a new operation token. This replaces the previous token while sharing source bytes and parsed messages. Saving is a separate destination-owner operation.

The adapter directly depends on `OfficeIMO.Core`, `OfficeIMO.IWork`, and `OfficeIMO.PowerPoint`. It does not add iWork support to the default PowerPoint package graph.

Font-relative character tracking is mapped to `PowerPointTextRun.CharacterSpacingPoints` when the effective font size is known. PPTX stores hundredths of a point and reports quantization. The qualified image deck now uses the default editable path; shared PDF/image rendering reports its current spacing approximation.

See the [iWork support matrix](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.iwork-support-matrix.md) for supported structures and conversion limits.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Convert | 0 | 1 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.PowerPoint.IWork` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
