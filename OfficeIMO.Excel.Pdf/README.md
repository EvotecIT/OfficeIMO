# OfficeIMO.Excel.Pdf - Excel to PDF export

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Excel.Pdf)](https://www.nuget.org/packages/OfficeIMO.Excel.Pdf)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Excel.Pdf?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Excel.Pdf)

`OfficeIMO.Excel.Pdf` exports `OfficeIMO.Excel` workbooks to PDF through the first-party `OfficeIMO.Pdf` engine. It also imports logical PDF tables into editable Excel worksheets.

## Install

```powershell
dotnet add package OfficeIMO.Excel.Pdf
```

## Quick start

```csharp
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;

using var workbook = ExcelDocument.Load("report.xlsx");
workbook.SaveAsPdf("report.pdf");
```

## Examples

### Export selected sheets with worksheet print settings

```csharp
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Pdf;

using var workbook = ExcelDocument.Load("monthly-report.xlsx");

var options = new ExcelPdfSaveOptions {
    SheetNames = new[] { "Summary", "Revenue", "Costs" },
    UseWorksheetPrintAreas = true,
    UseWorksheetPageSetup = true,
    UseWorksheetHeadersAndFooters = true,
    UseWorksheetPageBreaks = true,
    PageSize = PageSizes.A4.Landscape(),
    Margins = PageMargins.UniformCentimeters(1.2)
};

workbook.SaveAsPdf("monthly-report.pdf", options);
```

### Export a workbook to bytes or a stream

```csharp
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;

using var workbook = ExcelDocument.Load("statement.xlsx");

byte[] pdfBytes = workbook.ToPdf();

using var stream = File.Create("statement.pdf");
workbook.SaveAsPdf(stream);
```

### Surface mapping warnings

```csharp
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Pdf;

using var workbook = ExcelDocument.Load("dashboard.xlsx");
var options = new ExcelPdfSaveOptions {
    IncludeSheetHeadings = true,
    RespectWorksheetHiddenRowsAndColumns = true,
    UseWorksheetCharts = true
}.UseProfile(PdfExportProfile.Faithful);

options.TextFallbacks = PdfTextFallbackFeatures.Default;
options.AllowSystemFontEmbedding = true;

var result = workbook.TrySaveAsPdf("dashboard.pdf", options);
if (!result.Succeeded) {
    foreach (string diagnostic in result.Diagnostics) {
        Console.WriteLine(diagnostic);
    }
}

foreach (var warning in result.Warnings) {
    Console.WriteLine($"{warning.Source}: {warning.Message}");
}

result.Report.RequireNoErrorWarnings();
```

## Import PDF tables

```csharp
using OfficeIMO.Excel.Pdf;

var results = PdfExcelTableConverterExtensions.SavePdfTablesAsExcel(
    "statement.pdf",
    "statement-tables.xlsx");

foreach (var table in results) {
    Console.WriteLine($"{table.SheetName}: page {table.PageNumber}");
}
```

### Import only selected PDF pages

```csharp
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Pdf;

var results = PdfExcelTableConverterExtensions.SavePdfTablesAsExcel(
    "bank-statement.pdf",
    "bank-statement-q1.xlsx",
    new PdfExcelTableImportOptions {
        PageRanges = new[] { new PdfPageRange(1, 3) },
        MaxRows = 250
    });

Console.WriteLine($"Imported {results.Count} table(s).");
```

## What it maps

- Workbook sheets, selected sheet lists, visible used ranges, print areas, page setup, margins, orientation, and worksheet page breaks.
- Repeated print-title rows, headers, footers, page/date/time/sheet/workbook tokens, and supported header/footer images.
- Cell display values, common number formats, fills, font emphasis, alignment, borders, merged cells, links, row heights, column widths, conditional fills/data bars/icons, and table layout primitives.
- Supported worksheet images and common chart snapshots through shared OfficeIMO drawing primitives.
- Profile presets through `ExcelPdfSaveOptions.UseProfile(...)`, plus shared `TextFallbacks` and `AllowSystemFontEmbedding` controls for Unicode, symbols, and emoji.
- Per-operation conversion warnings through `PdfDocumentConversionResult.Report` or `PdfSaveResult.Report`.

## Boundaries

- Workbook reading stays in `OfficeIMO.Excel`.
- PDF layout and writing stay in `OfficeIMO.Pdf`.
- This package should remain a translation adapter, not a second PDF engine.
- Fidelity gaps should be documented as warnings or deeper current-state notes, not hidden in marketing text.

## Related packages

- [OfficeIMO.Excel](../OfficeIMO.Excel/README.md) - Excel workbook model.
- [OfficeIMO.Pdf](../OfficeIMO.Pdf/README.md) - PDF engine.
- [OfficeIMO.Html.Pdf](../OfficeIMO.Html.Pdf/README.md) - HTML/PDF bridge.

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`.
- License: MIT.
- Repository: [EvotecIT/OfficeIMO](https://github.com/EvotecIT/OfficeIMO)
