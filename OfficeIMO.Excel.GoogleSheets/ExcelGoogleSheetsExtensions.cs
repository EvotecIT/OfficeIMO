namespace OfficeIMO.Excel.GoogleSheets {
    /// <summary>
    /// Planning helpers for Excel to Google Sheets translation.
    /// </summary>
    public static class ExcelGoogleSheetsExtensions {
        private static readonly IGoogleSheetsExporter DefaultExporter = new GoogleSheetsExporter();

        public static GoogleSheetsTranslationPlan CreateGoogleSheetsTranslationPlan(
            this ExcelDocument document,
            GoogleSheetsSaveOptions? options = null) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            return DefaultExporter.BuildPlan(document, options);
        }

        public static GoogleSheetsBatch CreateGoogleSheetsBatch(
            this ExcelDocument document,
            GoogleSheetsSaveOptions? options = null) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            return DefaultExporter.BuildBatch(document, options);
        }

        public static Task<GoogleSpreadsheetReference> ExportToGoogleSheetsAsync(
            this ExcelDocument document,
            GoogleWorkspace.GoogleWorkspaceSession session,
            GoogleSheetsSaveOptions? options = null,
            CancellationToken cancellationToken = default) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            return DefaultExporter.ExportAsync(document, session, options, cancellationToken);
        }
    }
}
