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
    }
}
