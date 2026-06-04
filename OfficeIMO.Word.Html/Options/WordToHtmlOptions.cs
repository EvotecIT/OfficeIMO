namespace OfficeIMO.Word.Html {
    /// <summary>
    /// Options controlling Word to HTML conversion.
    /// </summary>
    public class WordToHtmlOptions {
        /// <summary>
        /// Optional font family applied to created runs during conversion.
        /// </summary>
        public string? FontFamily { get; set; }

        /// <summary>
        /// When true, includes run font information as inline styles.
        /// </summary>
        public bool IncludeFontStyles { get; set; }

        /// <summary>
        /// When set, includes list style information in generated HTML.
        /// </summary>
        public bool IncludeListStyles { get; set; }

        /// <summary>
        /// When true, emits reusable CSS classes and a head stylesheet for Word list definitions.
        /// Default is false to preserve the legacy inline/list-attribute output shape.
        /// </summary>
        public bool IncludeListDefinitions { get; set; }

        /// <summary>
        /// When true, paragraph styles are emitted as CSS classes.
        /// </summary>
        public bool IncludeParagraphClasses { get; set; }

        /// <summary>
        /// When true, run character styles are emitted as CSS classes.
        /// </summary>
        public bool IncludeRunClasses { get; set; }

        /// <summary>
        /// When true, includes run color information as inline styles.
        /// </summary>
        public bool IncludeRunColorStyles { get; set; }

        /// <summary>
        /// When true, includes run highlight information as inline styles.
        /// </summary>
        public bool IncludeRunHighlightStyles { get; set; }

        /// <summary>
        /// When true, includes paragraph spacing information as inline styles.
        /// </summary>
        public bool IncludeParagraphSpacingStyles { get; set; }

        /// <summary>
        /// When true, includes paragraph indentation information as inline styles.
        /// </summary>
        public bool IncludeParagraphIndentationStyles { get; set; }

        /// <summary>
        /// When true, footnotes are exported to HTML. Set to false to omit footnotes.
        /// </summary>
        public bool ExportFootnotes { get; set; } = true;

        /// <summary>
        /// When true, endnotes are exported to HTML. Set to false to omit endnotes.
        /// </summary>
        public bool ExportEndnotes { get; set; } = true;

        /// <summary>
        /// When true, Word comments are exported as linked HTML references and a comments section.
        /// Default is false so review metadata is not exposed unless requested.
        /// </summary>
        public bool ExportComments { get; set; }

        /// <summary>
        /// When true, Word section headers and footers are exported as semantic HTML
        /// <c>header</c> and <c>footer</c> regions with section/type metadata.
        /// Default is false to preserve the legacy body-only output.
        /// </summary>
        public bool ExportHeadersAndFooters { get; set; }

        /// <summary>
        /// When true, custom document properties are exported as typed HTML meta tags.
        /// Default is false so callers explicitly choose whether custom metadata is browser-visible.
        /// </summary>
        public bool IncludeCustomProperties { get; set; }

        /// <summary>
        /// When true, wraps exported document content in per-section <c>section</c>
        /// elements that preserve Word page size, orientation, and margin metadata.
        /// Default is false to preserve the legacy flat body output.
        /// </summary>
        public bool IncludeSectionMetadata { get; set; }

        /// <summary>
        /// When true, emits table column width metadata as HTML <c>colgroup</c>
        /// and <c>col</c> elements when the Word table exposes usable column widths.
        /// Default is false to preserve the legacy row-first table output.
        /// </summary>
        public bool IncludeTableColumnGroups { get; set; }

        /// <summary>
        /// When true (default), embeds images as base64 data URIs. When false,
        /// uses the image file paths instead.
        /// </summary>
        public bool EmbedImagesAsBase64 { get; set; } = true;

        /// <summary>
        /// Additional meta tags to include in the HTML head. Each tuple represents
        /// the <c>name</c> and <c>content</c> attributes of a meta element.
        /// </summary>
        public List<(string Name, string Content)> AdditionalMetaTags { get; } = new();

        /// <summary>
        /// Additional link tags to include in the HTML head. Each tuple represents
        /// the <c>rel</c> and <c>href</c> attributes of a link element.
        /// </summary>
        public List<(string Rel, string Href)> AdditionalLinkTags { get; } = new();

        /// <summary>
        /// When true, injects a small, built-in "Word-like" CSS into the HTML &lt;head&gt; to make output readable out-of-the-box.
        /// Default is false to preserve legacy behavior.
        /// </summary>
        public bool IncludeDefaultCss { get; set; } = false;
    }
}
