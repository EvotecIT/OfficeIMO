using OfficeIMO.Drawing;
using OfficeIMO.GoogleWorkspace;

namespace OfficeIMO.Excel.GoogleSheets {
    /// <summary>
    /// Options for importing a Google spreadsheet.
    /// </summary>
    public sealed class GoogleSheetsImportOptions {
        public const long DefaultMaxResponseBytes = 128L * 1024L * 1024L;
        public GoogleWorkspaceImportMode Mode { get; set; } = GoogleWorkspaceImportMode.DriveExport;
        public IReadOnlyList<string> Ranges { get; set; } = Array.Empty<string>();
        public string? Fields { get; set; }
        public ExcelLoadOptions LoadOptions { get; set; } = new ExcelLoadOptions {
            AccessMode = DocumentAccessMode.ReadWrite,
        };
        public IProgress<OfficeIMO.GoogleWorkspace.Drive.GoogleDriveTransferProgress>? Progress { get; set; }
        public long MaxResponseBytes { get; set; } = DefaultMaxResponseBytes;
        public int MaxSheets { get; set; } = 256;
        /// <summary>
        /// Maximum number of native cell values and row/column metadata entries that may be
        /// projected. Dimension metadata shares this budget because it can materialize row and
        /// column state even when the response contains no cell values.
        /// </summary>
        public long MaxCells { get; set; } = 1_000_000L;
        public long MaxDimensionGroupMembers { get; set; } = 1_000_000L;
    }

    /// <summary>
    /// Result of a Google Sheets import. The caller owns and must dispose <see cref="Document"/>.
    /// </summary>
    public sealed class GoogleSheetsImportResult {
        public GoogleSheetsImportResult(ExcelDocument document, GoogleSpreadsheetReference source, OfficeIMO.GoogleWorkspace.TranslationReport report) {
            Document = document ?? throw new ArgumentNullException(nameof(document));
            Source = source ?? throw new ArgumentNullException(nameof(source));
            Report = report ?? throw new ArgumentNullException(nameof(report));
        }

        public ExcelDocument Document { get; }
        public GoogleSpreadsheetReference Source { get; }
        public OfficeIMO.GoogleWorkspace.TranslationReport Report { get; }
    }
}
