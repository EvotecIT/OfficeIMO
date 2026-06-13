namespace OfficeIMO.Word.Markdown {
    /// <summary>
    /// Provides settings that control how a Word document is converted to Markdown.
    /// </summary>
    public class WordToMarkdownOptions {
        /// <summary>
        /// Font family whose runs should be rendered as inline code. When <c>null</c>,
        /// <see cref="FontResolver.Resolve(string)"/> is used with "monospace" to determine the code font.
        /// </summary>
        public string? FontFamily { get; set; }

        /// <summary>
        /// Enables wrapping underlined text with &lt;u&gt; tags.
        /// </summary>
        public bool EnableUnderline { get; set; }

        /// <summary>
        /// Enables wrapping highlighted text with == delimiters.
        /// </summary>
        public bool EnableHighlight { get; set; }

        /// <summary>
        /// Determines how images are exported during Markdown conversion.
        /// Default is <see cref="ImageExportMode.Base64"/>.
        /// </summary>
        public ImageExportMode ImageExportMode { get; set; } = ImageExportMode.Base64;

        /// <summary>
        /// When <see cref="ImageExportMode"/> is set to <see cref="ImageExportMode.File"/>,
        /// images are written to this directory. If not specified, the current working directory is used.
        /// </summary>
        public string? ImageDirectory { get; set; }

        /// <summary>
        /// Emits externally linked images as Markdown image references instead of failing extraction.
        /// This keeps existing documents with linked or cid-based images convertible even when the
        /// binary image payload is not stored in the package.
        /// </summary>
        public bool FallbackExternalImagesToLinks { get; set; } = true;

        /// <summary>
        /// Controls how Word page breaks are represented in Markdown.
        /// Default is <see cref="MarkdownPageBreakMode.SemanticBlock"/> so the break can round-trip.
        /// </summary>
        public MarkdownPageBreakMode PageBreakMode { get; set; } = MarkdownPageBreakMode.SemanticBlock;

        /// <summary>
        /// Controls how converter-detected Word content without a native Markdown representation is handled.
        /// Default is <see cref="MarkdownUnsupportedContentMode.WarnOnly"/>.
        /// </summary>
        public MarkdownUnsupportedContentMode UnsupportedContentMode { get; set; } = MarkdownUnsupportedContentMode.WarnOnly;

        /// <summary>
        /// Controls whether supported visual content, such as charts, is rendered as Markdown image fallbacks.
        /// Default is <see cref="MarkdownVisualFallbackMode.None"/> to keep Markdown semantic unless requested.
        /// </summary>
        public MarkdownVisualFallbackMode VisualFallbackMode { get; set; } = MarkdownVisualFallbackMode.None;

        /// <summary>
        /// Directory used for generated visual fallback resources when <see cref="VisualFallbackMode"/>
        /// is <see cref="MarkdownVisualFallbackMode.SvgFile"/>. When saving to a Markdown file and this
        /// is not set, resources are written to a sidecar folder next to the Markdown file.
        /// </summary>
        public string? VisualFallbackDirectory { get; set; }

        /// <summary>
        /// Optional path prefix used in Markdown image links for generated visual fallback resources.
        /// When saving to a Markdown file and this is not set, the sidecar folder name is used.
        /// </summary>
        public string? VisualFallbackPathPrefix { get; set; }

        /// <summary>
        /// Optional callback for non-fatal conversion warnings, such as external image fallback.
        /// </summary>
        public Action<string>? OnWarning { get; set; }

        /// <summary>
        /// Enables exporting section headers and footers as semantic fenced blocks instead of omitting them.
        /// The fenced payload contains markdown for the header/footer body and can be reparsed with
        /// <see cref="CreateReaderOptions(OfficeIMO.Markdown.MarkdownReaderOptions.MarkdownDialectProfile)"/>.
        /// </summary>
        public bool IncludeHeadersAndFootersAsSemanticBlocks { get; set; }

        /// <summary>
        /// Creates reader options that understand the Word-specific semantic fenced blocks emitted by this converter.
        /// </summary>
        /// <param name="profile">Base markdown reader profile to start from.</param>
        /// <returns>Configured reader options.</returns>
        public OfficeIMO.Markdown.MarkdownReaderOptions CreateReaderOptions(
            OfficeIMO.Markdown.MarkdownReaderOptions.MarkdownDialectProfile profile = OfficeIMO.Markdown.MarkdownReaderOptions.MarkdownDialectProfile.OfficeIMO) {
            var options = OfficeIMO.Markdown.MarkdownReaderOptions.CreateProfile(profile);
            WordMarkdownSemanticBlocks.ConfigureReaderOptions(options);
            return options;
        }
    }
}
