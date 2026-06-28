using OfficeIMO.Html;

namespace OfficeIMO.Word.Html {
    /// <summary>
    /// Options for saving an HTML-to-Word-to-HTML capability gallery scenario.
    /// </summary>
    public sealed class WordHtmlCapabilityGalleryOptions {
        /// <summary>Stable scenario identifier used for file names and manifests.</summary>
        public string ScenarioId { get; set; } = "word-html-roundtrip";

        /// <summary>Human-readable scenario title.</summary>
        public string Title { get; set; } = "Word HTML Roundtrip";

        /// <summary>Options used for importing source HTML into Word.</summary>
        public HtmlToWordOptions? ImportOptions { get; set; }

        /// <summary>Options used for exporting the Word document back to HTML.</summary>
        public WordToHtmlOptions? ExportOptions { get; set; }

        /// <summary>Resource URL policy used when building the source HTML resource manifest.</summary>
        public HtmlUrlPolicy? ResourceUrlPolicy { get; set; }
    }
}
