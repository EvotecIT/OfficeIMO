# OfficeIMO.Word.IWork

`OfficeIMO.Word.IWork` is the opt-in adapter for importing modern Apple Pages files into editable `OfficeIMO.Word` documents. Installing `OfficeIMO.Word` alone does not add the iWork reader.

```bash
dotnet add package OfficeIMO.Word.IWork
```

```csharp
using OfficeIMO.IWork;
using OfficeIMO.Word.IWork;

IWorkSourceDocument source = IWorkSourceDocument.Open("source.pages");
using PagesToWordResult result = source.ToWordDocumentResult(
    new IWorkConversionOptions { Mode = IWorkConversionMode.Auto });

Console.WriteLine(result.Report.ProjectionKind);
Console.WriteLine(result.HasLoss);
result.Value.Save("converted.docx");
```

`IWorkSourceDocument.Open` reads and bounds the source independently of destination policy. `ToWordDocument` returns the converted document directly; `ToWordDocumentResult` also exposes the typed Pages projection, diagnostics, preserved source records, and exact editable-versus-visual-fallback result. `WordIWorkConverter.ConvertPagesToWord*` provides equivalent path and stream convenience entry points.

Individual table row heights and column widths are carried into DOCX in twips under the adapter’s precision policy.

Choose the acceptance policy explicitly when source details cannot be represented:

```csharp
var options = new IWorkConversionOptions {
    Mode = IWorkConversionMode.Auto,
    AllowPartialEditableReconstruction = true,
    RequireCompleteVisualCoverage = true
};
```

This retains bounded recoverable editable content and reports incomplete details. If editable output cannot be produced, it rejects a first-page or composite preview. Both settings default to `false`. Inspect `Report.IsPartialEditableReconstruction` before accepting the output; `Report.RequireCompleteEditableReconstruction()` rejects explicitly partial reconstruction. `Report.RequireNoLoss()` also rejects unassessed record fidelity. These policies do not bypass source limits or destination safety checks.

Under the partial policy, positioned tables become flowing editable Word tables and finite measurements are rounded to DOCX units. `IWORK_PAGES_TABLE_LAYOUT_APPROXIMATED` and `IWORK_PAGES_DOCX_PRECISION` identify those approximations; original geometry remains on the source projection.

The path and stream convenience APIs accept cancellation after the options:

```csharp
using var cancellation = new CancellationTokenSource();
using PagesToWordResult cancellable = WordIWorkConverter.ConvertPagesToWordResult(
    "source.pages", readOptions: null, conversionOptions: options,
    cancellationToken: cancellation.Token);
```

This token governs loading, projection, and destination construction. It also governs later projections from `cancellable.Source`; reopen the source with a new token after cancellation. Saving is a separate destination-owner operation.

The adapter directly depends on `OfficeIMO.Core`, `OfficeIMO.IWork`, and `OfficeIMO.Word`. It does not add iWork support to the default Word package graph.

See the [iWork support matrix](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.iwork-support-matrix.md) for supported structures and conversion limits.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Convert | 0 | 1 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.Word.IWork` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
