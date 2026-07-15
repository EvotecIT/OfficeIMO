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
        public GoogleDocsTabOptions Tabs { get; set; } = new GoogleDocsTabOptions();
        public GoogleDocsReplaceOptions Replace { get; set; } = new GoogleDocsReplaceOptions();
        public GoogleDocsCommentMode Comments { get; set; } = GoogleDocsCommentMode.UnanchoredDriveComments;
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

    public enum GoogleDocsTabStrategy {
        FirstTab = 0,
        SelectedTab = 1,
        ReplaceEveryTab = 2,
    }

    /// <summary>Explicit tab selection for reads and writes.</summary>
    public sealed class GoogleDocsTabOptions {
        public GoogleDocsTabStrategy Strategy { get; set; } = GoogleDocsTabStrategy.FirstTab;
        public string? TabId { get; set; }
    }

    public enum GoogleDocsRevisionConflictMode {
        RequireRevision = 0,
        MergeAgainstTargetRevision = 1,
        OverwriteLatest = 2,
    }

    /// <summary>Collaboration policy for replacing an existing Google document.</summary>
    public sealed class GoogleDocsReplaceOptions {
        public GoogleDocsRevisionConflictMode ConflictMode { get; set; } = GoogleDocsRevisionConflictMode.RequireRevision;
        public string? ExpectedRevisionId { get; set; }
    }

    public enum GoogleDocsCommentMode {
        Skip = 0,
        UnanchoredDriveComments = 1,
    }
}
