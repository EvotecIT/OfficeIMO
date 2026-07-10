# OfficeIMO.Excel.OpenDocument

Explicit, dependency-free-at-runtime conversion between `OfficeIMO.Excel` workbooks and native `OfficeIMO.OpenDocument` spreadsheets.

```csharp
using OfficeIMO.Excel;
using OfficeIMO.Excel.OpenDocument;

using ExcelDocument workbook = ExcelDocument.Load("report.xlsx");
var conversion = workbook.ToOpenDocument();
conversion.Document.Save("report.ods");

foreach (var mapping in conversion.Report.Mappings) {
    Console.WriteLine($"{mapping.Feature}: {mapping.Status} ({mapping.Count})");
}
```

The adapter maps worksheets, typed cell values, formulas, hyperlinks, merges, row/column layout, named ranges, and a basic style subset. It reports approximated or unsupported workbook features instead of claiming silent full fidelity. ODS repeat expansion is bounded by `ExcelOpenDocumentConversionOptions`.
