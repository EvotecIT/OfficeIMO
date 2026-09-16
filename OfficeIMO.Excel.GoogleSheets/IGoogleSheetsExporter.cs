namespace OfficeIMO.Excel.GoogleSheets {
    /// <summary>
    /// Planning, compilation, and export contract for Excel to Google Sheets implementations.
    /// </summary>
    public interface IGoogleSheetsExporter {
        /// <summary>Reports source features and fidelity risks without contacting Google.</summary>
        GoogleSheetsTranslationPlan BuildPlan(ExcelDocument document, GoogleSheetsSaveOptions? options = null);
        /// <summary>Compiles a workbook into provider-neutral requests without contacting Google.</summary>
        GoogleSheetsBatch BuildBatch(ExcelDocument document, GoogleSheetsSaveOptions? options = null);
        /// <summary>Creates or replaces a Google spreadsheet through the supplied session.</summary>
        Task<GoogleSpreadsheetReference> ExportAsync(
            ExcelDocument document,
            GoogleWorkspace.GoogleWorkspaceSession session,
            GoogleSheetsSaveOptions? options = null,
            CancellationToken cancellationToken = default);
    }
}
