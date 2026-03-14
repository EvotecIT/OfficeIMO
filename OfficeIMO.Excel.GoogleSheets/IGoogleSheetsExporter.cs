namespace OfficeIMO.Excel.GoogleSheets {
    /// <summary>
    /// Future export contract for Excel to Google Sheets implementations.
    /// </summary>
    public interface IGoogleSheetsExporter {
        GoogleSheetsTranslationPlan BuildPlan(ExcelDocument document, GoogleSheetsSaveOptions? options = null);
        GoogleSheetsBatch BuildBatch(ExcelDocument document, GoogleSheetsSaveOptions? options = null);
        Task<GoogleSpreadsheetReference> ExportAsync(
            ExcelDocument document,
            GoogleWorkspace.GoogleWorkspaceSession session,
            GoogleSheetsSaveOptions? options = null,
            CancellationToken cancellationToken = default);
    }
}
