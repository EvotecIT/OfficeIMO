using OfficeIMO.GoogleWorkspace;

namespace OfficeIMO.Word.GoogleDocs {
    /// <summary>
    /// Planning-time options for Word to Google Docs export.
    /// </summary>
    public sealed class GoogleDocsSaveOptions {
        public GoogleDriveFileLocation Location { get; set; } = new GoogleDriveFileLocation();
        public string? Title { get; set; }
        public GoogleWorkspaceFidelityPolicy FidelityPolicy { get; set; } = new GoogleWorkspaceFidelityPolicy();
        public GoogleDocsUnsupportedFeatureOptions UnsupportedFeatures { get; set; } = new GoogleDocsUnsupportedFeatureOptions();
        public GoogleDocsInlineImageMode InlineImageMode { get; set; } = GoogleDocsInlineImageMode.Placeholder;
    }

    public sealed class GoogleDocsUnsupportedFeatureOptions {
        public UnsupportedFeatureMode FloatingContent { get; set; } = UnsupportedFeatureMode.WarnAndSkip;
        public UnsupportedFeatureMode Charts { get; set; } = UnsupportedFeatureMode.WarnAndSkip;
        public UnsupportedFeatureMode SmartArt { get; set; } = UnsupportedFeatureMode.WarnAndSkip;
        public UnsupportedFeatureMode ContentControls { get; set; } = UnsupportedFeatureMode.WarnAndSkip;
        public UnsupportedFeatureMode EmbeddedObjects { get; set; } = UnsupportedFeatureMode.WarnAndSkip;
        public UnsupportedFeatureMode Watermarks { get; set; } = UnsupportedFeatureMode.WarnAndSkip;
        public UnsupportedFeatureMode Comments { get; set; } = UnsupportedFeatureMode.WarnAndSkip;
        public UnsupportedFeatureMode Equations { get; set; } = UnsupportedFeatureMode.WarnAndSkip;
    }

    public enum GoogleDocsInlineImageMode {
        Placeholder = 0,
        TemporaryPublicDriveLease = 1,
    }
}
