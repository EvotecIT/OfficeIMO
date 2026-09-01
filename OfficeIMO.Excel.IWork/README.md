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

The adapter directly depends on `OfficeIMO.Core`, `OfficeIMO.IWork`, and `OfficeIMO.Excel`. It does not add iWork support to the default Excel package graph.

See the [iWork support matrix](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.iwork-support-matrix.md) for supported structures and conversion limits.
