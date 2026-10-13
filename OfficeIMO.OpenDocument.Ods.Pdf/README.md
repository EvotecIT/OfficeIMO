# OfficeIMO.OpenDocument.Ods.Pdf

Bidirectional ODS and PDF conversion without Word or PowerPoint dependencies. Forward `ToPdfDocumentResult()` calls expose the typed ODS-to-Excel report in `SourceConversionReports` and PDF-layout diagnostics in `Report`; `ConversionReports`, `HasLoss`, and `RequireNoLoss()` cover both stages. PDF to ODS reconstructs detected tables through the PDF-to-Excel and Excel-to-ODS adapters and reports omitted non-table page content.

PDF imports accept `PdfTablesToExcelOptions.PreserveColumnProportions` through `pdfOptions`. This optional policy retains detected relative column widths, with the [shared import settings and fallback behavior](../OfficeIMO.Excel.Pdf/README.md#import-structured-pdf-data). ODS widths use the spreadsheet adapter's approximate unit conversion; physical PDF dimensions and exact printer geometry are not guaranteed. `PdfOdsConversionReport.PdfReport.Entries` exposes whether each table applied the proportions.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Convert | 0 | 2 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.OpenDocument.Ods.Pdf` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
