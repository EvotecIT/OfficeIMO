using OfficeIMO.GoogleWorkspace;

namespace OfficeIMO.Excel.GoogleSheets {
    /// <summary>
    /// Result metadata for a created or updated Google Spreadsheet.
    /// </summary>
    public sealed class GoogleSpreadsheetReference : GoogleDriveFileReference {
        /// <summary>Gets or sets the Sheets spreadsheet identifier.</summary>
        public string? SpreadsheetId { get; set; }
        /// <summary>Gets or sets the observed Drive version, when available.</summary>
        public long? DriveVersion { get; set; }
        /// <summary>Gets or sets the observed modification time, when available.</summary>
        public DateTimeOffset? ModifiedTime { get; set; }
        /// <summary>Gets or sets translation notices associated with this reference.</summary>
        public TranslationReport Report { get; set; } = new TranslationReport();
    }
}
