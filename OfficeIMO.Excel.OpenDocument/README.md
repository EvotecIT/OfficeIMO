# OfficeIMO.Excel.OpenDocument

`OfficeIMO.Excel.OpenDocument` explicitly converts between `OfficeIMO.Excel` workbooks and native `OfficeIMO.OpenDocument` spreadsheets. It does not invoke Excel or LibreOffice; the adapter depends on the two OfficeIMO object-model packages it connects.

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

The adapter maps worksheets, typed cell values, formulas, hyperlinks, merges, row/column layout, named ranges, and a basic style subset. `ExcelOpenDocumentConversionOptions` bounds rows, columns, converted cells, and merge materialization in both directions. Content omitted by those limits or disabled style options is returned as a `Skipped` mapping rather than silently disappearing.

## Dependency footprint

- **External:** None.
- **OfficeIMO:** `OfficeIMO.Excel` and `OfficeIMO.OpenDocument`; the adapter owns bounded feature mapping and fidelity reports.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.
