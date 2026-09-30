# OfficeIMO.Excel.IWork

`OfficeIMO.Excel.IWork` is the opt-in adapter for importing modern Apple Numbers files into editable `OfficeIMO.Excel` workbooks. Installing `OfficeIMO.Excel` alone does not add the iWork reader.

```bash
dotnet add package OfficeIMO.Excel.IWork
```

```csharp
using OfficeIMO.Excel.IWork;
using OfficeIMO.IWork;

IWorkSourceDocument source = IWorkSourceDocument.Open("source.numbers");
using NumbersToExcelResult result = source.ToExcelDocumentResult(
    new IWorkConversionOptions { Mode = IWorkConversionMode.Auto });

Console.WriteLine(result.Report.ProjectionKind);
Console.WriteLine(result.HasLoss);
result.Value.Save("converted.xlsx");
```

`IWorkSourceDocument.Open` reads and bounds the source independently of destination policy. `ToExcelDocument` returns the converted workbook directly; `ToExcelDocumentResult` also exposes the typed Numbers projection, diagnostics, preserved source records, and exact editable-versus-visual-fallback result. `ExcelIWorkConverter.ConvertNumbersToExcel*` provides equivalent path and stream convenience entry points.

Choose the acceptance policy explicitly when source details cannot be represented:

```csharp
var options = new IWorkConversionOptions {
    Mode = IWorkConversionMode.Auto,
    AllowPartialEditableReconstruction = true,
    RequireCompleteVisualCoverage = true
};
```

This retains bounded recoverable editable content and reports incomplete details. If editable output cannot be produced, it rejects a first-page or composite preview. Both settings default to `false`. Inspect `Report.IsPartialEditableReconstruction` before accepting the output; `Report.RequireCompleteEditableReconstruction()` rejects explicitly partial reconstruction. `Report.RequireNoLoss()` also rejects unassessed record fidelity. These policies do not bypass source limits or destination safety checks.

Set `NormalizeWorksheetNames = true` to use the Excel owner's rules for invalid, long, or colliding worksheet names. The result's `WorksheetMappings` links each destination name to its one-based source sheet/table position and original name. Each source table still receives its own worksheet; table-local formulas retain their references. Renames produce `IWORK_NUMBERS_WORKSHEET_RENAMED`, an approximation diagnostic. This does not add cross-table formula support.

```csharp
using NumbersToExcelResult normalized = source.ToExcelDocumentResult(
    new IWorkConversionOptions { NormalizeWorksheetNames = true });
foreach (NumbersWorksheetMapping mapping in normalized.WorksheetMappings) {
    Console.WriteLine($"{mapping.SourceSheetIndex}/{mapping.SourceTableIndex}: {mapping.RequestedName} -> {mapping.DestinationName}");
}
```

The path and stream convenience APIs accept cancellation after the options:

```csharp
using var cancellation = new CancellationTokenSource();
using NumbersToExcelResult cancellable = ExcelIWorkConverter.ConvertNumbersToExcelResult(
    "source.numbers", readOptions: null, conversionOptions: options,
    cancellationToken: cancellation.Token);
```

This token governs loading, projection, and destination construction. It also governs later projections from `cancellable.Source`; reopen the source with a new token after cancellation. Saving is a separate destination-owner operation.

The adapter directly depends on `OfficeIMO.Core`, `OfficeIMO.IWork`, and `OfficeIMO.Excel`. It does not add iWork support to the default Excel package graph.

See the [iWork support matrix](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.iwork-support-matrix.md) for supported structures and conversion limits.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Convert | 0 | 1 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.Excel.IWork` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
