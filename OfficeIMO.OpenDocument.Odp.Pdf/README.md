# OfficeIMO.OpenDocument.Odp.Pdf

Bidirectional ODP and PDF conversion without Word or Excel dependencies. Forward `ToPdfDocumentResult()` calls expose the typed ODP-to-PowerPoint report in `SourceConversionReports` and PDF-layout diagnostics in `Report`; `ConversionReports`, `HasLoss`, and `RequireNoLoss()` cover both stages. PDF to ODP defaults to one rendered PDF page per slide for visual fidelity; use `PdfPowerPointImportOptions.CreateEditableTables()` when editable detected tables are the goal.
