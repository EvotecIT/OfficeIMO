# OfficeIMO.OpenDocument.Ods.Pdf

Bidirectional ODS and PDF conversion without Word or PowerPoint dependencies. Forward `ToPdfDocumentResult()` calls expose the typed ODS-to-Excel report in `SourceConversionReports` and PDF-layout diagnostics in `Report`; `ConversionReports`, `HasLoss`, and `RequireNoLoss()` cover both stages. PDF to ODS reconstructs detected tables through the PDF-to-Excel and Excel-to-ODS adapters and reports omitted non-table page content.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Convert | 0 | 2 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.OpenDocument.Ods.Pdf` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
