using OfficeIMO.GoogleWorkspace;

namespace OfficeIMO.Excel.GoogleSheets {
    /// <summary>
    /// Result metadata for a created or updated Google Spreadsheet.
    /// </summary>
    public sealed class GoogleSpreadsheetReference : GoogleDriveFileReference {
        public string? SpreadsheetId { get; set; }
        public TranslationReport Report { get; set; } = new TranslationReport();
    }
}
