namespace OfficeIMO.Excel.GoogleSheets {
    /// <summary>Imports a Google spreadsheet into a caller-owned Excel document.</summary>
    public interface IGoogleSheetsImporter {
        /// <summary>Loads a spreadsheet through Drive-exported XLSX or native Sheets projection.</summary>
        Task<GoogleSheetsImportResult> ImportAsync(
            string spreadsheetId,
            OfficeIMO.GoogleWorkspace.GoogleWorkspaceSession session,
            GoogleSheetsImportOptions? options = null,
            CancellationToken cancellationToken = default);
    }
}
