namespace OfficeIMO.Excel.GoogleSheets {
    public interface IGoogleSheetsImporter {
        Task<GoogleSheetsImportResult> ImportAsync(
            string spreadsheetId,
            OfficeIMO.GoogleWorkspace.GoogleWorkspaceSession session,
            GoogleSheetsImportOptions? options = null,
            CancellationToken cancellationToken = default);
    }
}
