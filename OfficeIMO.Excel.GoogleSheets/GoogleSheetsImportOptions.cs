using OfficeIMO.Drawing;

namespace OfficeIMO.Excel.GoogleSheets {
    /// <summary>
    /// Selects the Google Sheets import path.
    /// </summary>
    public enum GoogleSheetsImportMode {
        /// <summary>Export the Google-native spreadsheet to XLSX through Drive and load it with OfficeIMO.</summary>
        DriveExport,
        /// <summary>Read the native Sheets resource and project the supported model into OfficeIMO.</summary>
        Native,
    }

    /// <summary>
    /// Options for importing a Google spreadsheet.
    /// </summary>
    public sealed class GoogleSheetsImportOptions {
        public const long DefaultMaxResponseBytes = 128L * 1024L * 1024L;
        public GoogleSheetsImportMode Mode { get; set; } = GoogleSheetsImportMode.DriveExport;
        public IReadOnlyList<string> Ranges { get; set; } = Array.Empty<string>();
        public string? Fields { get; set; }
        public ExcelLoadOptions LoadOptions { get; set; } = new ExcelLoadOptions {
            AccessMode = DocumentAccessMode.ReadWrite,
        };
        public IProgress<OfficeIMO.GoogleWorkspace.Drive.GoogleDriveTransferProgress>? Progress { get; set; }
        public long MaxResponseBytes { get; set; } = DefaultMaxResponseBytes;
        public int MaxSheets { get; set; } = 256;
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
