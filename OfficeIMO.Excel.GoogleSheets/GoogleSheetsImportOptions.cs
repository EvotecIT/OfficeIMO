using OfficeIMO.Drawing;
using OfficeIMO.GoogleWorkspace;

namespace OfficeIMO.Excel.GoogleSheets {
    /// <summary>
    /// Options for importing a Google spreadsheet.
    /// </summary>
    public sealed class GoogleSheetsImportOptions {
        /// <summary>Default upper bound for Drive-exported XLSX or a native Sheets API response, in bytes.</summary>
        public const long DefaultMaxResponseBytes = 128L * 1024L * 1024L;
        /// <summary>Gets or sets Drive-export or native API import; defaults to Drive export.</summary>
        public GoogleWorkspaceImportMode Mode { get; set; } = GoogleWorkspaceImportMode.DriveExport;
        /// <summary>Gets or sets optional A1 ranges for native import.</summary>
        public IReadOnlyList<string> Ranges { get; set; } = Array.Empty<string>();
        /// <summary>Gets or sets an optional Sheets API field mask for native import.</summary>
        public string? Fields { get; set; }
        /// <summary>Gets or sets options used to load Drive-exported XLSX; defaults to read-write access.</summary>
        public ExcelLoadOptions LoadOptions { get; set; } = new ExcelLoadOptions {
            AccessMode = DocumentAccessMode.ReadWrite,
        };
        /// <summary>Gets or sets an optional observer for Drive-export transfer progress.</summary>
        public IProgress<OfficeIMO.GoogleWorkspace.Drive.GoogleDriveTransferProgress>? Progress { get; set; }
        /// <summary>Gets or sets the positive byte limit for Drive-exported XLSX or a native API response.</summary>
        public long MaxResponseBytes { get; set; } = DefaultMaxResponseBytes;
        /// <summary>Gets or sets the positive maximum number of sheets projected by native import.</summary>
        public int MaxSheets { get; set; } = 256;
        /// <summary>
        /// Maximum number of native cell values and row/column metadata entries that may be
        /// projected. Dimension metadata shares this budget because it can materialize row and
        /// column state even when the response contains no cell values.
        /// </summary>
        public long MaxCells { get; set; } = 1_000_000L;
        /// <summary>Gets or sets the positive limit on native dimension-group members.</summary>
        public long MaxDimensionGroupMembers { get; set; } = 1_000_000L;
    }

    /// <summary>
    /// Result of a Google Sheets import. The caller owns and must dispose <see cref="Document"/>.
    /// </summary>
    public sealed class GoogleSheetsImportResult {
        /// <summary>Creates an import result from a caller-owned document, source metadata, and notices.</summary>
        public GoogleSheetsImportResult(ExcelDocument document, GoogleSpreadsheetReference source, OfficeIMO.GoogleWorkspace.TranslationReport report) {
            Document = document ?? throw new ArgumentNullException(nameof(document));
            Source = source ?? throw new ArgumentNullException(nameof(source));
            Report = report ?? throw new ArgumentNullException(nameof(report));
        }

        /// <summary>Gets the imported Excel document, which the caller must dispose.</summary>
        public ExcelDocument Document { get; }
        /// <summary>Gets Google source metadata, including the observed Drive version.</summary>
        public GoogleSpreadsheetReference Source { get; }
        /// <summary>Gets fidelity and operation notices from import.</summary>
        public OfficeIMO.GoogleWorkspace.TranslationReport Report { get; }
    }
}
