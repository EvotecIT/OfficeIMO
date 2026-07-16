namespace OfficeIMO.Word.GoogleDocs {
    /// <summary>Support level for one Word/Google Docs translation capability.</summary>
    public enum GoogleDocsFeatureSupportLevel {
        Native = 0,
        Flattened = 1,
        Partial = 2,
        DriveFallback = 3,
        Unsupported = 4,
    }

    /// <summary>Code-owned support-matrix row used by documentation and preflight tooling.</summary>
    public sealed class GoogleDocsFeatureSupport {
        public GoogleDocsFeatureSupport(string feature, GoogleDocsFeatureSupportLevel export, GoogleDocsFeatureSupportLevel import, string notes) {
            Feature = feature;
            Export = export;
            Import = import;
            Notes = notes;
        }

        public string Feature { get; }
        public GoogleDocsFeatureSupportLevel Export { get; }
        public GoogleDocsFeatureSupportLevel Import { get; }
        public string Notes { get; }
    }

    /// <summary>Authoritative feature matrix for the current package version.</summary>
    public static class GoogleDocsFeatureSupportCatalog {
        private static readonly IReadOnlyList<GoogleDocsFeatureSupport> FeaturesValue = new[] {
            new GoogleDocsFeatureSupport("Paragraphs, runs, headings, lists and hyperlinks", GoogleDocsFeatureSupportLevel.Native, GoogleDocsFeatureSupportLevel.Native, "Native import projects core styles; Drive DOCX export remains the broad-fidelity fallback."),
            new GoogleDocsFeatureSupport("Tables and merged cells", GoogleDocsFeatureSupportLevel.Native, GoogleDocsFeatureSupportLevel.Partial, "Export replays table structure and supported styling; native import projects simple cells."),
            new GoogleDocsFeatureSupport("Headers, footers, footnotes and bookmarks", GoogleDocsFeatureSupportLevel.Native, GoogleDocsFeatureSupportLevel.DriveFallback, "Native import reports these segments and directs callers to Drive-export import for exact placement."),
            new GoogleDocsFeatureSupport("Document tabs", GoogleDocsFeatureSupportLevel.Native, GoogleDocsFeatureSupportLevel.Native, "Reads are tab-aware; callers select one tab or deliberately replace/flatten every tab."),
            new GoogleDocsFeatureSupport("Comments", GoogleDocsFeatureSupportLevel.Flattened, GoogleDocsFeatureSupportLevel.DriveFallback, "Word comments become unanchored Drive comments with author context and replies."),
            new GoogleDocsFeatureSupport("Inline images", GoogleDocsFeatureSupportLevel.Partial, GoogleDocsFeatureSupportLevel.DriveFallback, "Export supports placeholders or explicit temporary public Drive leases; Drive export preserves imported binaries."),
            new GoogleDocsFeatureSupport("Page and section layout", GoogleDocsFeatureSupportLevel.Partial, GoogleDocsFeatureSupportLevel.DriveFallback, "Supported page size, margins, headers and footers are native; columns and Word-only pagination are reported."),
            new GoogleDocsFeatureSupport("All-caps and tab leaders", GoogleDocsFeatureSupportLevel.Flattened, GoogleDocsFeatureSupportLevel.DriveFallback, "All-caps is materialized in text and tab leaders are emitted as characters so appearance survives."),
            new GoogleDocsFeatureSupport("Charts, SmartArt, floating content and embedded objects", GoogleDocsFeatureSupportLevel.Unsupported, GoogleDocsFeatureSupportLevel.DriveFallback, "Caller fidelity policy controls fail/skip behavior; Drive DOCX import is the broad read fallback."),
            new GoogleDocsFeatureSupport("Equations, watermarks and content controls", GoogleDocsFeatureSupportLevel.Unsupported, GoogleDocsFeatureSupportLevel.DriveFallback, "These features are diagnosed explicitly and are not silently inferred."),
        };

        public static IReadOnlyList<GoogleDocsFeatureSupport> Features => FeaturesValue;
    }
}
