namespace OfficeIMO.Word.Markdown {
    /// <summary>
    /// Provides settings that control how a Word document is converted to Markdown.
    /// </summary>
    public class WordMarkdownExportOptions {
        internal WordMarkdownExportOptions CloneForConversion() => (WordMarkdownExportOptions)MemberwiseClone();

        /// <summary>Rejects fidelity loss before returning the target document or writing its destination.</summary>
        public bool FailOnLoss { get; set; }

        internal Action<WordMarkdownConversionDiagnostic>? DiagnosticSink { get; set; }

        /// <summary>
        /// Font family whose runs should be rendered as inline code. When <c>null</c>,
        /// <see cref="WordFontResolver.Resolve(string)"/> is used with "monospace" to determine the code font.
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
        /// Default is <see cref="WordMarkdownImageExportMode.Base64"/>. Saving to a path without options writes images
        /// as files next to the Markdown file.
        /// </summary>
        public WordMarkdownImageExportMode ImageExportMode { get; set; } = WordMarkdownImageExportMode.Base64;

        /// <summary>
        /// When <see cref="WordMarkdownImageExportMode"/> is set to <see cref="WordMarkdownImageExportMode.File"/>,
        /// images are written to this directory. If not specified, the current working directory is used.
        /// </summary>
        public string? ImageDirectory { get; set; }

        /// <summary>
        /// Maximum number of image bytes that may be materialized for one base64 data URI.
        /// The default is 32 MiB. Set a different positive value for trusted documents that
        /// intentionally contain larger embedded images.
        /// </summary>
        public long MaxEmbeddedImageBytes { get; set; } = 32L * 1024L * 1024L;

        /// <summary>
        /// Maximum Word list nesting level materialized during conversion. The default is 128.
        /// </summary>
        public int MaxListNestingDepth { get; set; } = 128;

        /// <summary>
        /// Emits externally linked images as Markdown image references instead of failing extraction.
        /// This keeps existing documents with linked or cid-based images convertible even when the
        /// binary image payload is not stored in the package.
        /// </summary>
        public bool FallbackExternalImagesToLinks { get; set; } = true;

        /// <summary>
        /// Controls how Word page breaks are represented in Markdown.
        /// Default is <see cref="WordMarkdownPageBreakMode.SemanticBlock"/> so the break can round-trip.
        /// </summary>
        public WordMarkdownPageBreakMode PageBreakMode { get; set; } = WordMarkdownPageBreakMode.SemanticBlock;

        /// <summary>
        /// Controls how converter-detected Word content without a native Markdown representation is handled.
        /// Default is <see cref="WordMarkdownUnsupportedContentMode.WarnOnly"/>.
        /// </summary>
        public WordMarkdownUnsupportedContentMode UnsupportedContentMode { get; set; } = WordMarkdownUnsupportedContentMode.WarnOnly;

        /// <summary>
        /// Controls whether supported visual content, such as charts, is rendered as Markdown image fallbacks.
        /// Default is <see cref="WordMarkdownVisualFallbackMode.None"/> to keep Markdown semantic unless requested.
        /// </summary>
        public WordMarkdownVisualFallbackMode VisualFallbackMode { get; set; } = WordMarkdownVisualFallbackMode.None;

        /// <summary>
        /// Directory used for generated visual fallback resources when <see cref="VisualFallbackMode"/>
        /// is <see cref="WordMarkdownVisualFallbackMode.SvgFile"/>. When saving to a Markdown file and this
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
        /// <see cref="CreateReaderOptions(OfficeIMO.Markdown.MarkdownLoadOptions.MarkdownDialectProfile)"/>.
        /// </summary>
        public bool IncludeHeadersAndFootersAsSemanticBlocks { get; set; }

        /// <summary>
        /// Creates reader options that understand the Word-specific semantic fenced blocks emitted by this converter.
        /// </summary>
        /// <param name="profile">Base markdown reader profile to start from.</param>
        /// <returns>Configured reader options.</returns>
        public OfficeIMO.Markdown.MarkdownLoadOptions CreateReaderOptions(
            OfficeIMO.Markdown.MarkdownLoadOptions.MarkdownDialectProfile profile = OfficeIMO.Markdown.MarkdownLoadOptions.MarkdownDialectProfile.OfficeIMO) {
            var options = OfficeIMO.Markdown.MarkdownLoadOptions.CreateProfile(profile);
            WordMarkdownSemanticBlocks.ConfigureReaderOptions(options);
            return options;
        }
    }
}
