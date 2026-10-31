# OfficeIMO.Pdf.Excel

Reconstruct Excel content from supported PDF pages.

```sh
dotnet add package OfficeIMO.Pdf.Excel
```

## Import structured PDF data

```csharp
using OfficeIMO;
using OfficeIMO.Pdf.Excel;
using OfficeIMO.Pdf;

PdfDocument pdf = PdfDocument.Load("statement.pdf");
OfficeConversionReport report = pdf.SaveAsExcel("statement-tables.xlsx");

foreach (OfficeDiagnostic finding in report.Diagnostics) {
    Console.WriteLine($"{finding.Code}: {finding.Message}");
}

Console.WriteLine($"Non-table page content detected: {report.HasLoss}");
```

### Import only selected PDF pages

```csharp
using OfficeIMO;
using OfficeIMO.Pdf.Excel;
using OfficeIMO.Pdf;

PdfDocument pdf = PdfDocument.Load("bank-statement.pdf");
OfficeConversionReport report = pdf.SaveAsExcel(
    "bank-statement-q1.xlsx",
    new ExcelPdfImportOptions {
        ReadOptions = new PdfLoadOptions {
            PageSelection = PdfPageSelection.Parse("1-3")
        },
        MaxRows = 250
    });

Console.WriteLine($"Imported {report.ItemsWritten} table(s).");
report.ThrowIfLoss(); // rejects row truncation and non-table page content
```

`ReadOptions` also exposes the canonical Fast/Structured profile and semantic work limits for large or deliberately bounded imports. It is ignored when the source is already a `PdfDocumentReadResult`.

Tables without explicit header evidence retain all source rows, with generated Excel column names. If you have reviewed the source and confirmed its first row as headers, set `UseFirstRowAsHeader = true`. This applies to each logical table with unknown schema; tables with an established header keep their existing body boundary. The import excludes the confirmed header from type detection and `MaxRows`. Recognition results from [standalone images and scanned PDFs](../OfficeIMO.Pdf.Ocr/README.md) use this same importer.

Set `PreserveColumnProportions = true` to retain the detected table's relative column widths, including quarter-turn tables. With the default `AutoFitColumns = true`, auto-fit supplies the total width and detected proportions determine each column's share. With `AutoFitColumns = false`, the total comes from worksheet default widths. If a mapped column would exceed Excel's 255-character width limit, one common scale reduces the total while retaining the proportions. The option is off by default. Missing or invalid geometry keeps the existing widths; the reconstructed worksheet retains the mapped widths. These widths use Excel character units and can leave individual values unfitted. They do not guarantee physical PDF dimensions, authored ruling boundaries, or exact printer geometry. PDF-to-ODS uses the same policy through the shared importer, with approximate spreadsheet unit conversion.

Compatible table segments continue across adjacent pages by default. The shared PDF table analysis classifies Boolean, percentage, date-time, time-only, currency, numeric, and text columns with confidence; the import options decide which detected families become typed Excel cells. Consistent currency columns become decimal cells and retain their detected symbol or ISO 4217 code, prefix or suffix placement, source spacing, and each cell's visible fractional precision in the Excel number format. Mixed currency tokens or affix conventions remain text so the import does not silently normalize their meaning.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Convert | 0 | 1 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.Pdf.Excel` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
