# OfficeIMO.OpenDocument.Ods.Pdf

Bidirectional ODS and PDF conversion without Word or PowerPoint dependencies. Forward `ToPdfDocumentResult()` calls expose the typed ODS-to-Excel report in `SourceConversionReports` and PDF-layout diagnostics in `Report`; `ConversionReports`, `HasLoss`, and `RequireNoLoss()` cover both stages. PDF to ODS reconstructs detected tables through the PDF-to-Excel and Excel-to-ODS adapters and reports omitted non-table page content.
