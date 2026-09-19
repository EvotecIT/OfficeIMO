namespace OfficeIMO.PowerPoint.GoogleSlides {
    /// <summary>How a presentation feature is handled in one translation direction.</summary>
    public enum GoogleSlidesFeatureSupportLevel {
        /// <summary>Represented with editable native Google Slides or OfficeIMO content.</summary>
        Native = 0,
        /// <summary>Preserved visually as an image rather than as editable content.</summary>
        Rasterized = 1,
        /// <summary>Only the documented subset maps directly.</summary>
        Partial = 2,
        /// <summary>Preserved through Drive-exported PPTX import rather than native projection.</summary>
        DriveFallback = 3,
        /// <summary>No supported translation is currently provided.</summary>
        Unsupported = 4,
    }
    /// <summary>Export and import support levels for one presentation feature group.</summary>
    public sealed class GoogleSlidesFeatureSupport {
        /// <summary>Creates a feature entry with direction-specific support and caveats.</summary>
        public GoogleSlidesFeatureSupport(string feature, GoogleSlidesFeatureSupportLevel export, GoogleSlidesFeatureSupportLevel import, string notes) { Feature = feature; Export = export; Import = import; Notes = notes; }
        /// <summary>Gets the described presentation feature group.</summary>
        public string Feature { get; }
        /// <summary>Gets the support level when exporting to Google Slides.</summary>
        public GoogleSlidesFeatureSupportLevel Export { get; }
        /// <summary>Gets the support level when importing into OfficeIMO.</summary>
        public GoogleSlidesFeatureSupportLevel Import { get; }
        /// <summary>Gets the feature-specific limitation or behavior.</summary>
        public string Notes { get; }
    }
    /// <summary>Code-owned matrix of major Google Slides translation capabilities.</summary>
    public static class GoogleSlidesFeatureSupportCatalog {
        private static readonly IReadOnlyList<GoogleSlidesFeatureSupport> FeaturesValue = new[] {
            new GoogleSlidesFeatureSupport("Slides, ordering, size and solid backgrounds", GoogleSlidesFeatureSupportLevel.Partial, GoogleSlidesFeatureSupportLevel.Native, "Export scales and centers source coordinates when the Google page size differs; Slides does not expose page-size updates. Deterministic page IDs are scoped to one apply and are not synchronization identities."),
            new GoogleSlidesFeatureSupport("Text boxes, run typography and hyperlinks", GoogleSlidesFeatureSupportLevel.Native, GoogleSlidesFeatureSupportLevel.Native, "Font family, size, color, bold, italic, single underline, strike, small caps, superscript and subscript map natively. All-caps is materialized; Office-only underline and double-strike variants use their closest Google appearance."),
            new GoogleSlidesFeatureSupport("Tables", GoogleSlidesFeatureSupportLevel.Partial, GoogleSlidesFeatureSupportLevel.Partial, "Unmerged table geometry and cell text map natively; merged cells trigger whole-slide fallback by default, while advanced borders and themes are not fully projected. Drive-exported PPTX remains the broad import fallback."),
            new GoogleSlidesFeatureSupport("Pictures", GoogleSlidesFeatureSupportLevel.Partial, GoogleSlidesFeatureSupportLevel.Partial, "Uncropped PNG, JPEG, and GIF pictures export as native image objects through temporary public Drive leases; cleanup is attempted after export and failures are reported. Cropped or unsupported-format pictures render with whole-slide fallback by default, or are skipped under PreferNativeAndReport. Native import skips an image it cannot download and reports a warning; callers can separately choose Drive-exported PPTX import for broader fidelity."),
            new GoogleSlidesFeatureSupport("Basic shapes", GoogleSlidesFeatureSupportLevel.Partial, GoogleSlidesFeatureSupportLevel.Partial, "Common geometry maps natively; unsupported custom geometry uses whole-slide fallback by default or is skipped under PreferNativeAndReport."),
            new GoogleSlidesFeatureSupport("Speaker notes", GoogleSlidesFeatureSupportLevel.Native, GoogleSlidesFeatureSupportLevel.Native, "Only the speaker-notes BODY placeholder is writable in the Slides API."),
            new GoogleSlidesFeatureSupport("Charts and SmartArt", GoogleSlidesFeatureSupportLevel.Rasterized, GoogleSlidesFeatureSupportLevel.DriveFallback, "PowerPoint charts are not equivalent to linked Google Sheets charts; complex slides use renderer-owned PNG fallback."),
            new GoogleSlidesFeatureSupport("Video and audio", GoogleSlidesFeatureSupportLevel.Rasterized, GoogleSlidesFeatureSupportLevel.DriveFallback, "Media is rendered; no matching editable source-link contract is currently defined."),
            new GoogleSlidesFeatureSupport("Masters, themes and layouts", GoogleSlidesFeatureSupportLevel.Partial, GoogleSlidesFeatureSupportLevel.DriveFallback, "Blank-slide authoring and template-copy workflows are supported; full master mutation is intentionally not inferred."),
            new GoogleSlidesFeatureSupport("Transitions, animations, diagrams, equations and OLE", GoogleSlidesFeatureSupportLevel.Rasterized, GoogleSlidesFeatureSupportLevel.DriveFallback, "These remain explicit fidelity boundaries and are preserved visually through complex-slide rendering."),
        };
        /// <summary>Gets the feature groups and their import/export support levels.</summary>
        public static IReadOnlyList<GoogleSlidesFeatureSupport> Features => FeaturesValue;
    }
}
