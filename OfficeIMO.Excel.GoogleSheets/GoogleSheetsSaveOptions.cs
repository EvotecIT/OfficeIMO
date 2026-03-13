using OfficeIMO.GoogleWorkspace;

namespace OfficeIMO.Excel.GoogleSheets {
    /// <summary>
    /// Planning-time options for Excel to Google Sheets export.
    /// </summary>
    public sealed class GoogleSheetsSaveOptions {
        public GoogleDriveFileLocation Location { get; set; } = new GoogleDriveFileLocation();
        public string? Title { get; set; }
        public bool IncludeCharts { get; set; } = true;
        public bool IncludePivotTables { get; set; } = true;
        public bool IncludeHeaderFooterMetadata { get; set; } = true;
        public bool PreserveUnsupportedFormulasAsText { get; set; }
        public bool TreatPrintLayoutAsDiagnosticOnly { get; set; } = true;
    }
}
