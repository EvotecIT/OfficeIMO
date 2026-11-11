# OfficeIMO.OpenDocument.Ods.Pdf

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.OpenDocument.Ods.Pdf)](https://www.nuget.org/packages/OfficeIMO.OpenDocument.Ods.Pdf)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.OpenDocument.Ods.Pdf?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.OpenDocument.Ods.Pdf)

Bidirectional ODS and PDF conversion without Word or PowerPoint dependencies. ODS to PDF converts the spreadsheet through [OfficeIMO.Excel.OpenDocument](../OfficeIMO.Excel.OpenDocument/README.md) and prints it with the [Excel PDF engine](../OfficeIMO.Excel.Pdf/README.md). PDF to ODS reconstructs detected tables through the PDF-to-Excel and Excel-to-ODS adapters.

## Install

```powershell
dotnet add package OfficeIMO.OpenDocument.Ods.Pdf
```

## Quick start

```csharp
using OfficeIMO;
using OfficeIMO.Excel.OpenDocument;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.OpenDocument;
using OfficeIMO.OpenDocument.Ods.Pdf;
using OfficeIMO.Pdf;

OdsDocument workbook = OdsDocument.Load("budget.ods");
OfficeConversionReport report = workbook.SaveAsPdf("budget.pdf", new OdsPdfExportOptions {
    PdfOptions = new ExcelPdfExportOptions { PageSize = PdfPageSizes.A4 }
});
foreach (OfficeDiagnostic finding in report.Diagnostics)
    Console.WriteLine($"{finding.Location}: {finding.LossKind} {finding.Message}");

// Reject the conversion when the spreadsheet stage skips or cannot map content.
workbook.SaveAsPdf("budget-strict.pdf", new OdsPdfExportOptions {
    ConversionOptions = new ExcelOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported }
});

PdfDocument pdf = PdfDocument.Load("statement.pdf");
OfficeConversionReport odsReport = pdf.SaveAsOds("statement.ods");
```

Forward conversion takes `OdsPdfExportOptions` (`ConversionOptions` for the OpenDocument import, `PdfOptions` for the PDF engine, `FailOnLoss`); `ToPdf(options, out report)`, `ToPdfAsync`, `SaveAsPdf` and `SaveAsPdfAsync` return the report of both stages. `FailOnLoss` throws `OfficeConversionException` when either stage reports loss; `ConversionOptions.LossPolicy` throws `OdfConversionLossException` from the spreadsheet stage. Reverse conversion takes `PdfOdsExportOptions` (`PdfOptions`, `OpenDocumentOptions`, `FailOnLoss`) on a `PdfDocument` or `PdfDocumentReadResult`: `ToOds(options, out report)`, `ToOdsAsync`, `SaveAsOds` and `SaveAsOdsAsync`. The report from `ToOds(options, out report)` and `ToOdsAsync` is a `PdfOdsConversionReport`; cast it to read the separate PDF and OpenDocument stage reports. `SaveAsOds` and `SaveAsOdsAsync` return a plain `OfficeConversionReport` with the diagnostics of both stages.

## What it exports

- Each visible sheet with content prints through the Excel PDF engine, with searchable cell text, cell styles, merges, column widths and row heights.
- Sheet print ranges (`OdsSheet.PrintRanges`) on the same sheet become the print area.
- Charts, pivot tables, validations and conditional formatting follow the [ODS-to-Excel mapping](../OfficeIMO.Excel.OpenDocument/README.md); content outside that mapping is reported, not dropped silently.
- The blank styled run that LibreOffice and Excel write to pad each row to the sheet edge does not become printed pages.

PDF imports accept `PdfExcelExportOptions.PreserveColumnProportions` through `PdfOdsExportOptions.PdfOptions`. This optional policy retains detected relative column widths, with the [shared import settings and fallback behavior](../OfficeIMO.Excel.Pdf/README.md#import-structured-pdf-data). ODS widths use the spreadsheet adapter's approximate unit conversion; physical PDF dimensions and exact printer geometry are not guaranteed. `PdfOdsConversionReport.PdfReport.Entries` exposes whether each table applied the proportions.

## Current limits

- The spreadsheet stage inherits the ODS-to-XLSX limits: borders, value-dependent alignment, blank cells styled only by row, column or family defaults, and unsupported validation, chart and pivot shapes are reported as loss.
- Page breaks and scaling follow the Excel PDF print layout, not LibreOffice or Excel.
- PDF to ODS reconstructs table-shaped content only; free text, charts, formulas and scans without OCR are not rebuilt.

## Related packages

- [OfficeIMO.OpenDocument](../OfficeIMO.OpenDocument/README.md) reads, edits and writes ODS.
- [OfficeIMO.Excel.OpenDocument](../OfficeIMO.Excel.OpenDocument/README.md) converts between ODS and XLSX.
- [OfficeIMO.OpenDocument.Odt.Pdf](../OfficeIMO.OpenDocument.Odt.Pdf/README.md) and [OfficeIMO.OpenDocument.Odp.Pdf](../OfficeIMO.OpenDocument.Odp.Pdf/README.md) cover ODT and ODP.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Convert | 1 | 1 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.OpenDocument.Ods.Pdf` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
