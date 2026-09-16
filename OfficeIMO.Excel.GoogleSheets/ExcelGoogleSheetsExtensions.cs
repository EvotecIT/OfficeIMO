namespace OfficeIMO.Excel.GoogleSheets {
    /// <summary>
    /// Planning helpers for Excel to Google Sheets translation.
    /// </summary>
    public static class ExcelGoogleSheetsExtensions {
        private static readonly IGoogleSheetsExporter DefaultExporter = new GoogleSheetsExporter();
        private static readonly IGoogleSheetsImporter DefaultImporter = new GoogleSheetsImporter();

        /// <summary>Reports workbook features and fidelity risks without contacting Google.</summary>
        public static GoogleSheetsTranslationPlan BuildGoogleSheetsPlan(
            this ExcelDocument document,
            GoogleSheetsSaveOptions? options = null) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            return DefaultExporter.BuildPlan(document, options);
        }

        /// <summary>Compiles workbook content into provider-neutral Sheets requests without contacting Google.</summary>
        public static GoogleSheetsBatch BuildGoogleSheetsBatch(
            this ExcelDocument document,
            GoogleSheetsSaveOptions? options = null) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            return DefaultExporter.BuildBatch(document, options);
        }

        /// <summary>Creates or replaces a Google spreadsheet using the supplied session.</summary>
        public static Task<GoogleSpreadsheetReference> ExportToGoogleSheetsAsync(
            this ExcelDocument document,
            GoogleWorkspace.GoogleWorkspaceSession session,
            GoogleSheetsSaveOptions? options = null,
            CancellationToken cancellationToken = default) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            return DefaultExporter.ExportAsync(document, session, options, cancellationToken);
        }

        /// <summary>Imports a Google spreadsheet and returns a caller-owned Excel document.</summary>
        public static Task<GoogleSheetsImportResult> ImportGoogleSheetAsync(
            this GoogleWorkspace.GoogleWorkspaceSession session,
            string spreadsheetId,
            GoogleSheetsImportOptions? options = null,
            CancellationToken cancellationToken = default) {
            if (session == null) throw new ArgumentNullException(nameof(session));
            return DefaultImporter.ImportAsync(spreadsheetId, session, options, cancellationToken);
        }
    }
}
