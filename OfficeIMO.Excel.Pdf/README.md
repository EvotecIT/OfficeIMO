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

var options = new ExcelToPdfOptions {
    SheetNames = new[] { "Summary", "Revenue", "Costs" },
    UseWorksheetPrintAreas = true,
    UseWorksheetPageSetup = true,
    UseWorksheetHeadersAndFooters = true,
    UseWorksheetPageBreaks = true,
    UseWorksheetPrintTitleColumns = true,
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

byte[] pdfBytes = workbook.ToPdfBytes();

using var stream = File.Create("statement.pdf");
workbook.SaveAsPdf(stream);
```

### Surface mapping warnings

```csharp
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Pdf;

using var workbook = ExcelDocument.Load("dashboard.xlsx");
var options = new ExcelToPdfOptions {
    IncludeSheetHeadings = true,
    RespectWorksheetHiddenRowsAndColumns = true,
    UseWorksheetCharts = true
}.UseProfile(PdfExportProfile.Faithful);

options.TextFallbacks = PdfTextFallbackFeatures.Default;
options.ResourcePolicy = PdfResourcePolicy.CreateTrustedHost();

var result = workbook.SaveAsPdfResult("dashboard.pdf", options);
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

## Import structured PDF data

```csharp
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Pdf;

PdfDocument pdf = PdfDocument.Load("statement.pdf");
PdfExcelTableImportReport report = pdf.SaveTablesAsExcel("statement-tables.xlsx").RequireSuccess().Report!;

foreach (var table in report.Entries) {
    Console.WriteLine($"{table.SheetName}: page {table.PageNumber}");
}

Console.WriteLine($"Non-table page content detected: {report.HasOmittedPageContent}");
```

### Import only selected PDF pages

```csharp
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Pdf;

PdfDocument pdf = PdfDocument.Load("bank-statement.pdf");
PdfExcelTableImportReport report = pdf.SaveTablesAsExcel(
    "bank-statement-q1.xlsx",
    new PdfTablesToExcelOptions {
        ReadOptions = new PdfReadOptions {
            PageSelection = PdfPageSelection.Parse("1-3")
        },
        MaxRows = 250
    }).RequireSuccess().Report!;

Console.WriteLine($"Imported {report.Entries.Count} table(s).");
report.RequireNoLoss(); // rejects row truncation and non-table page content
```

`ReadOptions` also exposes the canonical Fast/Structured profile and semantic work limits for large or deliberately bounded imports. It is ignored when the source is already a `PdfDocumentReadResult`.

Compatible table segments continue across adjacent pages by default. The shared PDF table analysis classifies Boolean, percentage, date-time, time-only, currency, numeric, and text columns with confidence; the import options decide which detected families become typed Excel cells. Consistent currency columns become decimal cells and retain their detected symbol or ISO 4217 code, prefix or suffix placement, source spacing, and each cell's visible fractional precision in the Excel number format. Mixed currency tokens or affix conventions remain text so the import does not silently normalize their meaning.

## What it maps

- Workbook sheets, selected sheet lists, visible used ranges, print areas, page setup, margins, orientation, and worksheet page breaks.
- Repeated print-title rows and columns, headers, footers, page/date/time/sheet/workbook tokens, and supported header/footer images.
- Cell display values, common number formats, fills, font emphasis, alignment, borders, merged cells, links, row heights, column widths, conditional fills/data bars/icons, and table layout primitives. General alignment follows stored cell types: numeric and date values align right, Booleans align center, and text retains its text alignment. Boolean display uses `TRUE` and `FALSE`.
- Numeric cell text uses the Excel owner's display formatter, including optional decimal placeholders, single-digit scientific mantissas with exponent sign/case/padding, percentage scaling, grouping, negative sections, quoted/escaped numeric literals, literal percent signs, and the `[Red]` format color. Stored numeric precision is preserved before formatting, including values immediately beside a fraction midpoint. Other format colors are not projected. Imported formats use at most four sections; extra sections do not affect display selection.
- Text cells and text formula caches retain their text and font color, even when the text looks numeric and the cell has a numeric, date, or elapsed-time format.
- Supported worksheet images and common chart snapshots through shared OfficeIMO drawing primitives.
- Deterministic profile presets through `ExcelToPdfOptions.UseProfile(...)`. Applying a profile always resets the complete profile-owned option set, so reusing an options instance is history-independent.
- Shared `TextFallbacks` and `ResourcePolicy` controls for Unicode, symbols, emoji, and host-resource trust. The balanced default uses installed fonts while denying arbitrary local and remote reads; portable deterministic mode is explicit.
- Source-faithful zero-options output: worksheet-name headings are opt-in through `IncludeSheetHeadings`.
- Per-operation conversion warnings through `PdfDocumentConversionResult.Report` or `PdfSaveResult.Report`.

Worksheet print areas are exported separately, in their stored order, with each area starting on a new page. Both worksheet layouts honor local A1 cells and ranges, whole-row and whole-column selections, and multiple areas. Content outside the selected areas is excluded. Repeated title rows, first/even/odd headers and footers, and page numbering belong to the worksheet and continue across its areas. Internal links to a cell appearing in several areas target its first exported occurrence. Invalid or external print-area references are rejected instead of widening the selection to the used range.

`WorksheetCanvas` paginates using worksheet row heights and column widths, repeats configured title columns on horizontal pages, and applies the worksheet's down-then-over or over-then-down page order. Title columns left of a print area are included without exporting the intervening columns. A title column repeats after the page sequence reaches it; it is not moved onto an earlier page. Titles already inside a page's body are emitted once; titles beyond the print area's last column are excluded. Manual breaks remain on axes with an unlimited fit count, while a constrained fit axis determines its own page boundaries. Unspecified cell vertical alignment uses Excel's bottom alignment; explicit top and center settings take precedence.

`FlowTable` reflows cells into PDF tables and repeats title columns across requested horizontal chunks. Fit-to-height scales cell fonts, padding, and minimum row heights together. Rows can grow to retain their laid-out text, so this mode does not promise worksheet page counts.

Print-title rows can start inside the print area. Canvas automatic pages and both layouts' manual row chunks repeat only title rows reached before the current body segment, including a break within a title block. Flow-table automatic continuation still repeats leading header rows; title rows starting later in a single flowing table are not repeated automatically.

## Current limits

- Workbook content is read through `OfficeIMO.Excel`; layout and PDF writing use `OfficeIMO.Pdf`.
- Worksheet column widths and fit scaling use layout estimates. Exact Excel printer geometry and derived fit percentages are not guaranteed; fit-to-width documents can have different scale and page counts. Use producer comparisons for workflows that depend on identical pagination.
- PDF import is structured-data recovery. It reconstructs detected tables as worksheets; `SourceScope` and `HasOmittedPageContent` report text, source vector graphics, images, links, forms, annotations, or actions that are not represented by those tables. `HasLoss` and `RequireNoLoss()` include both that omitted page content and table-row truncation.
- The current reverse route recovers detected tables and structured values; arbitrary PDF page art is reported rather than presented as an editable workbook. Open recovery work is tracked in the repository [roadmap](../Docs/ROADMAP.md).

## Related packages

- [OfficeIMO.Excel](../OfficeIMO.Excel/README.md) - Excel workbook model.
- [OfficeIMO.Pdf](../OfficeIMO.Pdf/README.md) - PDF engine.
- [OfficeIMO.Html.Pdf](../OfficeIMO.Html.Pdf/README.md) - HTML/PDF bridge.

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`.
- License: MIT.
- Repository: [EvotecIT/OfficeIMO](https://github.com/EvotecIT/OfficeIMO)

## Dependency footprint

- **External:** None beyond the dependencies of its OfficeIMO format packages; no native or commercial PDF renderer.
- **OfficeIMO:** `OfficeIMO.Excel`, `OfficeIMO.Pdf`, and `OfficeIMO.Core` own layout mapping, rendering, table recovery, and reports.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Convert | 1 | 1 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.Excel.Pdf` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
