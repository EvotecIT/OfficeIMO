namespace OfficeIMO.PowerPoint.GoogleSlides {
    public enum GoogleSlidesFeatureSupportLevel { Native = 0, Rasterized = 1, Partial = 2, DriveFallback = 3, Unsupported = 4 }
    public sealed class GoogleSlidesFeatureSupport {
        public GoogleSlidesFeatureSupport(string feature, GoogleSlidesFeatureSupportLevel export, GoogleSlidesFeatureSupportLevel import, string notes) { Feature = feature; Export = export; Import = import; Notes = notes; }
        public string Feature { get; }
        public GoogleSlidesFeatureSupportLevel Export { get; }
        public GoogleSlidesFeatureSupportLevel Import { get; }
        public string Notes { get; }
    }
    public static class GoogleSlidesFeatureSupportCatalog {
        private static readonly IReadOnlyList<GoogleSlidesFeatureSupport> FeaturesValue = new[] {
            new GoogleSlidesFeatureSupport("Slides, ordering, size and solid backgrounds", GoogleSlidesFeatureSupportLevel.Native, GoogleSlidesFeatureSupportLevel.Native, "Deterministic page IDs are scoped to one apply and are not synchronization identities."),
            new GoogleSlidesFeatureSupport("Text boxes, core run styles and hyperlinks", GoogleSlidesFeatureSupportLevel.Native, GoogleSlidesFeatureSupportLevel.Native, "The first native style run is used as the editable baseline; Drive PPTX import preserves richer run structure."),
            new GoogleSlidesFeatureSupport("Tables", GoogleSlidesFeatureSupportLevel.Native, GoogleSlidesFeatureSupportLevel.Native, "Core geometry and cell text are native; advanced borders, merges, and themes use the Drive fallback."),
            new GoogleSlidesFeatureSupport("Pictures", GoogleSlidesFeatureSupportLevel.Native, GoogleSlidesFeatureSupportLevel.Native, "Export uses short-lived public Drive leases because Slides fetches images by URL."),
            new GoogleSlidesFeatureSupport("Basic shapes", GoogleSlidesFeatureSupportLevel.Native, GoogleSlidesFeatureSupportLevel.Partial, "Common geometry maps natively; exact custom geometry is rendered when complex-slide fallback is enabled."),
            new GoogleSlidesFeatureSupport("Speaker notes", GoogleSlidesFeatureSupportLevel.Native, GoogleSlidesFeatureSupportLevel.Native, "Only the speaker-notes BODY placeholder is writable in the Slides API."),
            new GoogleSlidesFeatureSupport("Charts and SmartArt", GoogleSlidesFeatureSupportLevel.Rasterized, GoogleSlidesFeatureSupportLevel.DriveFallback, "PowerPoint charts are not equivalent to linked Google Sheets charts; complex slides use renderer-owned PNG fallback."),
            new GoogleSlidesFeatureSupport("Video and audio", GoogleSlidesFeatureSupportLevel.Rasterized, GoogleSlidesFeatureSupportLevel.DriveFallback, "Media is rendered unless a future source-link contract proves matching semantics."),
            new GoogleSlidesFeatureSupport("Masters, themes and layouts", GoogleSlidesFeatureSupportLevel.Partial, GoogleSlidesFeatureSupportLevel.DriveFallback, "Blank-slide authoring and template-copy workflows are supported; full master mutation is intentionally not inferred."),
            new GoogleSlidesFeatureSupport("Transitions, animations, diagrams, equations and OLE", GoogleSlidesFeatureSupportLevel.Rasterized, GoogleSlidesFeatureSupportLevel.DriveFallback, "These remain explicit fidelity boundaries and are preserved visually through complex-slide rendering."),
        };
        public static IReadOnlyList<GoogleSlidesFeatureSupport> Features => FeaturesValue;
    }
}
