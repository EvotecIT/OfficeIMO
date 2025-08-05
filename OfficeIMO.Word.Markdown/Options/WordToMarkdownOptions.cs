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
    }
}
