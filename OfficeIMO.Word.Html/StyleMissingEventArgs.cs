namespace OfficeIMO.Word.Html {
    /// <summary>
    /// Event arguments for missing CSS class mapping during HTML-to-Word conversion.
    /// </summary>
    public class StyleMissingEventArgs : EventArgs {
        /// <summary>
        /// Creates a new instance.
        /// </summary>
        /// <param name="paragraph">Paragraph where the style was referenced.</param>
        /// <param name="className">The missing CSS class name.</param>
        public StyleMissingEventArgs(WordParagraph paragraph, string className) {
            Paragraph = paragraph ?? throw new ArgumentNullException(nameof(paragraph));
            ClassName = className ?? throw new ArgumentNullException(nameof(className));
        }

        /// <summary>
        /// Paragraph where the missing style occurred.
        /// </summary>
        public WordParagraph Paragraph { get; }

        /// <summary>
        /// Missing CSS class name.
        /// </summary>
        public string ClassName { get; }

        /// <summary>
        /// Optional mapping to a built-in paragraph style.
        /// </summary>
        public WordParagraphStyles? Style { get; set; }

        /// <summary>
        /// Optional mapping to a style ID.
        /// </summary>
        public string? StyleId { get; set; }
    }
}
