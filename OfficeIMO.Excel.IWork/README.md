# OfficeIMO.Excel.IWork

`OfficeIMO.Excel.IWork` is the opt-in adapter for importing modern Apple Numbers files into editable `OfficeIMO.Excel` workbooks. Installing `OfficeIMO.Excel` alone does not add the iWork reader.

```bash
dotnet add package OfficeIMO.Excel.IWork
```

```csharp
using OfficeIMO.Excel.IWork;
using OfficeIMO.IWork;

using IWorkNumbersLoadResult result = ExcelIWorkConverter.LoadNumbersWithReport(
    "source.numbers",
    new IWorkReadOptions { ImportMode = IWorkImportMode.Auto });

Console.WriteLine(result.ImportReport.ProjectionKind);
result.Document.Save("converted.xlsx");
```

`LoadNumbers` returns the editable Excel document directly. `LoadNumbersWithReport` also retains the bounded source model, typed Numbers projection, diagnostics, unsupported records, and exact editable-versus-visual-fallback result.

The adapter depends on `OfficeIMO.IWork` and `OfficeIMO.Excel`. It does not add iWork support to the default Excel package graph.

See the [iWork support matrix](../Docs/officeimo.iwork-support-matrix.md) for supported structures and conversion limits.
