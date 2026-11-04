# OfficeIMO.OpenDocument.Odp.Pdf

Bidirectional ODP and PDF conversion without Word or Excel dependencies. Forward conversion takes `OdpPdfExportOptions` (`ConversionOptions` for the OpenDocument import, `PdfOptions` for the PDF engine, `FailOnLoss`); `ToPdf(options, out report)`, `ToPdfAsync`, `SaveAsPdf` and `SaveAsPdfAsync` return the report of both stages. Reverse conversion takes `PdfOdpExportOptions` (`PdfOptions`, `OpenDocumentOptions`, `FailOnLoss`) on a `PdfDocument` or `PdfDocumentReadResult`: `ToOdp(options, out report)`, `ToOdpAsync`, `SaveAsOdp` and `SaveAsOdpAsync`. The reverse report is a `PdfOdpConversionReport` with the PDF and OpenDocument stage reports. PDF to ODP defaults to reconstructing supported editable content through the shared PowerPoint projection. An already reduced `PdfDocumentReadResult` resolves the default `Auto` profile to detected tables only. Use `PdfPowerPointExportOptions.CreateVisualPages()` for one rendered page image per slide, `CreateHybrid()` for visual pages with editable table overlays, or `CreateEditableTables()` for detected tables only.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Convert | 0 | 2 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.OpenDocument.Odp.Pdf` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
