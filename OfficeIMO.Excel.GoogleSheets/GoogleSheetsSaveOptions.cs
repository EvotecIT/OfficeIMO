using OfficeIMO.GoogleWorkspace;

namespace OfficeIMO.Excel.GoogleSheets {
    /// <summary>
    /// Planning-time options for Excel to Google Sheets export.
    /// </summary>
    public sealed class GoogleSheetsSaveOptions {
        public GoogleDriveFileLocation Location { get; set; } = new GoogleDriveFileLocation();
        public string? Title { get; set; }
        public GoogleWorkspaceFidelityPolicy FidelityPolicy { get; set; } = new GoogleWorkspaceFidelityPolicy();
        public GoogleSheetsUnsupportedFeatureOptions UnsupportedFeatures { get; set; } = new GoogleSheetsUnsupportedFeatureOptions();
    }

    public sealed class GoogleSheetsUnsupportedFeatureOptions {
        public UnsupportedFeatureMode Charts { get; set; } = UnsupportedFeatureMode.WarnAndSkip;
        public UnsupportedFeatureMode PivotTables { get; set; } = UnsupportedFeatureMode.WarnAndSkip;
        public UnsupportedFeatureMode PrintLayout { get; set; } = UnsupportedFeatureMode.WarnAndSkip;
    }
}
