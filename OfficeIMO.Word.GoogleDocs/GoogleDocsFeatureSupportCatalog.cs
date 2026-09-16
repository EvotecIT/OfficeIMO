namespace OfficeIMO.Word.GoogleDocs {
    /// <summary>Support level for one Word/Google Docs translation capability.</summary>
    public enum GoogleDocsFeatureSupportLevel {
        /// <summary>Represented as native editable content.</summary>
        Native = 0,
        /// <summary>Retained in a simpler target representation.</summary>
        Flattened = 1,
        /// <summary>Only the documented subset maps directly.</summary>
        Partial = 2,
        /// <summary>Recovered through Drive-exported DOCX rather than native projection.</summary>
        DriveFallback = 3,
        /// <summary>No supported direct translation is provided.</summary>
        Unsupported = 4,
    }

    /// <summary>Code-owned support-matrix row used by documentation and preflight tooling.</summary>
    public sealed class GoogleDocsFeatureSupport {
        /// <summary>Creates a feature row with direction-specific support and limitations.</summary>
        public GoogleDocsFeatureSupport(string feature, GoogleDocsFeatureSupportLevel export, GoogleDocsFeatureSupportLevel import, string notes) {
            Feature = feature;
            Export = export;
            Import = import;
            Notes = notes;
        }

        /// <summary>Gets the feature group described by this row.</summary>
        public string Feature { get; }
        /// <summary>Gets its support level when exporting to Google Docs.</summary>
        public GoogleDocsFeatureSupportLevel Export { get; }
        /// <summary>Gets its support level when importing into OfficeIMO.</summary>
        public GoogleDocsFeatureSupportLevel Import { get; }
        /// <summary>Gets feature-specific behavior and limitations.</summary>
        public string Notes { get; }
    }

    /// <summary>Authoritative feature matrix for the current package version.</summary>
    public static class GoogleDocsFeatureSupportCatalog {
        private static readonly IReadOnlyList<GoogleDocsFeatureSupport> FeaturesValue = new[] {
            new GoogleDocsFeatureSupport("Paragraphs, runs, headings, lists and hyperlinks", GoogleDocsFeatureSupportLevel.Partial, GoogleDocsFeatureSupportLevel.Partial, "Export emits external URI links but does not apply internal-anchor navigation; core text and heading styles are native. Native import projects core text and heading styles, flattens list markers, and reports hyperlinks without reconstructing them. Drive-exported DOCX remains the broad-fidelity import choice."),
            new GoogleDocsFeatureSupport("Tables and merged cells", GoogleDocsFeatureSupportLevel.Partial, GoogleDocsFeatureSupportLevel.Partial, "Export creates editable tables and supported cell merges, but nested tables inside cells are not represented by the Word inspection snapshot. Native import projects simple cells."),
            new GoogleDocsFeatureSupport("Headers, footers, footnotes and bookmarks", GoogleDocsFeatureSupportLevel.Partial, GoogleDocsFeatureSupportLevel.DriveFallback, "Export creates default header/footer segments and supported footnotes. A paragraph bookmark name becomes a Docs named range, not an exact Word point or range; additional bookmarks in the same paragraph are not preserved. First-page and even-page header/footer variants are skipped, as are footnotes inside header/footer segments. Native import reports segments without reconstructing their placement; use Drive-exported DOCX for broader import fidelity."),
            new GoogleDocsFeatureSupport("Document tabs", GoogleDocsFeatureSupportLevel.Partial, GoogleDocsFeatureSupportLevel.Flattened, "Creation writes to the default tab; replacement targets the first or selected tab, or explicitly replaces every tab. Export does not create a Word-authored tab hierarchy. Native import selects one tab or combines multiple tabs with headings in one Word document, without preserving tab structure."),
            new GoogleDocsFeatureSupport("Comments", GoogleDocsFeatureSupportLevel.Flattened, GoogleDocsFeatureSupportLevel.DriveFallback, "Word comments become unanchored Drive comments with author context and replies."),
            new GoogleDocsFeatureSupport("Inline images", GoogleDocsFeatureSupportLevel.Partial, GoogleDocsFeatureSupportLevel.DriveFallback, "Export uses readable placeholders by default. Opt-in native insertion supports PNG, JPEG, and GIF through a public Drive lease with no automatic expiry; the file can remain public if cleanup fails. Only the first positioned image in each Word run reaches the batch. Drive-exported DOCX preserves imported image binaries."),
            new GoogleDocsFeatureSupport("Page and section layout", GoogleDocsFeatureSupportLevel.Partial, GoogleDocsFeatureSupportLevel.DriveFallback, "Supported columns map to native section styles. The first section's paper size becomes document-wide; later sections can retain orientation and margins, but not a different paper size. Word-only pagination and unsupported header/footer variants are reported."),
            new GoogleDocsFeatureSupport("All-caps and tab leaders", GoogleDocsFeatureSupportLevel.Flattened, GoogleDocsFeatureSupportLevel.DriveFallback, "All-caps is materialized in text; tab leaders become fill characters rather than editable tab-leader metadata."),
            new GoogleDocsFeatureSupport("Charts, SmartArt, floating content and embedded objects", GoogleDocsFeatureSupportLevel.Unsupported, GoogleDocsFeatureSupportLevel.DriveFallback, "Caller fidelity policy controls fail/skip behavior; Drive DOCX import is the broad read fallback."),
            new GoogleDocsFeatureSupport("Equations, watermarks and content controls", GoogleDocsFeatureSupportLevel.Unsupported, GoogleDocsFeatureSupportLevel.DriveFallback, "No direct editable export is provided. Preflight counts top-level body equations, but equations in table cells or header/footer stories can be omitted without a notice. The caller's unsupported-feature policy governs features detected by preflight."),
        };

        /// <summary>Gets the current feature groups and direction-specific support levels.</summary>
        public static IReadOnlyList<GoogleDocsFeatureSupport> Features => FeaturesValue;
    }
}
