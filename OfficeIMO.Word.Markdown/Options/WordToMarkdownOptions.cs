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
