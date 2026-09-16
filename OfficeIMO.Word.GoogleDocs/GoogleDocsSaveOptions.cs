using OfficeIMO.GoogleWorkspace;
using OfficeIMO.Word;

namespace OfficeIMO.Word.GoogleDocs {
    /// <summary>
    /// Planning-time options for Word to Google Docs export.
    /// </summary>
    public sealed class GoogleDocsSaveOptions {
        /// <summary>Gets or sets the Drive destination or existing document target.</summary>
        public GoogleDriveFileLocation Location { get; set; } = new GoogleDriveFileLocation();
        /// <summary>Gets or sets an optional document title; blank uses the source title, file name, or <c>Document</c>.</summary>
        public string? Title { get; set; }
        /// <summary>Gets or sets the policy checked against fidelity notices before mutation.</summary>
        public GoogleWorkspaceFidelityPolicy FidelityPolicy { get; set; } = new GoogleWorkspaceFidelityPolicy();
        /// <summary>Gets or sets handling for Word features without dependable native Docs equivalents.</summary>
        public GoogleDocsUnsupportedFeatureOptions UnsupportedFeatures { get; set; } = new GoogleDocsUnsupportedFeatureOptions();
        /// <summary>Gets or sets how inline images are represented; defaults to placeholders.</summary>
        public GoogleDocsInlineImageMode InlineImageMode { get; set; } = GoogleDocsInlineImageMode.Placeholder;
        /// <summary>Gets or sets the tab selection used for reads and replacement.</summary>
        public GoogleDocsTabOptions Tabs { get; set; } = new GoogleDocsTabOptions();
        /// <summary>Gets or sets the revision policy for replacing an existing document.</summary>
        public GoogleDocsReplaceOptions Replace { get; set; } = new GoogleDocsReplaceOptions();
        /// <summary>Gets or sets how Word comments are transferred; defaults to unanchored Drive comments.</summary>
        public GoogleDocsCommentMode Comments { get; set; } = GoogleDocsCommentMode.UnanchoredDriveComments;
        /// <summary>
        /// Bounded renderer settings used when unsupported Word content is rasterized into fallback pages.
        /// The page range is always the complete document; output-count, pixel, byte, codec, and policy
        /// settings are honored.
        /// </summary>
        public WordImageExportOptions RasterFallbackImageOptions { get; set; } =
            CreateRasterFallbackImageOptions();

        private static WordImageExportOptions CreateRasterFallbackImageOptions() =>
            new WordImageExportOptions {
                MaximumOutputCount = 100,
                MaximumRasterPixels = 25_000_000,
                MaximumTotalRasterPixels = 250_000_000,
                MaximumTotalEncodedBytes = 128L * 1024 * 1024,
                MaximumDegreeOfParallelism = 1,
            };
    }

    /// <summary>Per-feature choices when a Word feature cannot be mapped faithfully to native Docs content.</summary>
    public sealed class GoogleDocsUnsupportedFeatureOptions {
        /// <summary>Gets or sets handling for floating content; defaults to warning and skip.</summary>
        public UnsupportedFeatureMode FloatingContent { get; set; } = UnsupportedFeatureMode.WarnAndSkip;
        /// <summary>Gets or sets handling for charts; defaults to warning and skip.</summary>
        public UnsupportedFeatureMode Charts { get; set; } = UnsupportedFeatureMode.WarnAndSkip;
        /// <summary>Gets or sets handling for SmartArt; defaults to warning and skip.</summary>
        public UnsupportedFeatureMode SmartArt { get; set; } = UnsupportedFeatureMode.WarnAndSkip;
        /// <summary>Gets or sets handling for content controls; defaults to warning and skip.</summary>
        public UnsupportedFeatureMode ContentControls { get; set; } = UnsupportedFeatureMode.WarnAndSkip;
        /// <summary>Gets or sets handling for embedded objects; defaults to warning and skip.</summary>
        public UnsupportedFeatureMode EmbeddedObjects { get; set; } = UnsupportedFeatureMode.WarnAndSkip;
        /// <summary>Gets or sets handling for watermarks; defaults to warning and skip.</summary>
        public UnsupportedFeatureMode Watermarks { get; set; } = UnsupportedFeatureMode.WarnAndSkip;
        /// <summary>Gets or sets handling for comments that cannot be represented as selected; defaults to warning and skip.</summary>
        public UnsupportedFeatureMode Comments { get; set; } = UnsupportedFeatureMode.WarnAndSkip;
        /// <summary>Gets or sets handling for equations; defaults to warning and skip.</summary>
        public UnsupportedFeatureMode Equations { get; set; } = UnsupportedFeatureMode.WarnAndSkip;
    }

    /// <summary>How exporter writes inline images into a Google document.</summary>
    public enum GoogleDocsInlineImageMode {
        /// <summary>Insert a textual placeholder instead of publishing image content.</summary>
        Placeholder = 0,
        /// <summary>Use short-lived public Drive image leases so Docs can fetch the content.</summary>
        TemporaryPublicDriveLease = 1,
    }

    /// <summary>Which document tab is targeted by an export or replacement.</summary>
    public enum GoogleDocsTabStrategy {
        /// <summary>Target the first document tab.</summary>
        FirstTab = 0,
        /// <summary>Target the tab identified by <see cref="GoogleDocsTabOptions.TabId"/>.</summary>
        SelectedTab = 1,
        /// <summary>Replace content in every document tab.</summary>
        ReplaceEveryTab = 2,
    }

    /// <summary>Explicit tab selection for reads and writes.</summary>
    public sealed class GoogleDocsTabOptions {
        /// <summary>Gets or sets the tab targeting strategy; defaults to first tab.</summary>
        public GoogleDocsTabStrategy Strategy { get; set; } = GoogleDocsTabStrategy.FirstTab;
        /// <summary>Gets or sets the required tab identifier for selected-tab targeting.</summary>
        public string? TabId { get; set; }
    }

    /// <summary>Revision behavior when replacing an existing Google document.</summary>
    public enum GoogleDocsRevisionConflictMode {
        /// <summary>Require the observed revision to match before each guarded write.</summary>
        RequireRevision = 0,
        /// <summary>Use the observed revision as a Docs target revision, allowing API-managed merge when possible.</summary>
        MergeAgainstTargetRevision = 1,
        /// <summary>Write against the latest remote content without a revision guard.</summary>
        OverwriteLatest = 2,
    }

    /// <summary>Collaboration policy for replacing an existing Google document.</summary>
    public sealed class GoogleDocsReplaceOptions {
        /// <summary>Gets or sets the conflict mode; defaults to requiring an exact revision.</summary>
        public GoogleDocsRevisionConflictMode ConflictMode { get; set; } = GoogleDocsRevisionConflictMode.RequireRevision;
        /// <summary>Gets or sets the revision observed during an earlier read or import.</summary>
        public string? ExpectedRevisionId { get; set; }
    }

    /// <summary>How source comments are handled during Docs export.</summary>
    public enum GoogleDocsCommentMode {
        /// <summary>Do not create Drive comments from Word comments.</summary>
        Skip = 0,
        /// <summary>Create unanchored Drive comments with available author and reply context.</summary>
        UnanchoredDriveComments = 1,
    }
}
