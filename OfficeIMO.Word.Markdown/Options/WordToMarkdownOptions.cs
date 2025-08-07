using OfficeIMO.Word;

namespace OfficeIMO.Word.Markdown {
    /// <summary>
    /// Options for Word to Markdown conversion.
    /// </summary>
    public class WordToMarkdownOptions {
        /// <summary>
        /// Optional font family applied to created runs during conversion.
        /// </summary>
        public string FontFamily { get; set; }

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
        public string ImageDirectory { get; set; }
    }
}
