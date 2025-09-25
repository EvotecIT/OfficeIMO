using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Provides serialization helpers for Word elements.
    /// </summary>
    public static class WordXmlExtensions {
        /// <summary>
        /// Returns the raw Open XML string representing the paragraph.
        /// </summary>
        /// <param name="paragraph">Paragraph to convert.</param>
        /// <returns>Outer XML of the underlying Open XML element.</returns>
        public static string ToXml(this WordParagraph paragraph) {
            if (paragraph == null) throw new ArgumentNullException(nameof(paragraph));
            return paragraph._paragraph?.OuterXml ?? string.Empty;
        }

        /// <summary>
        /// Creates a <see cref="WordParagraph"/> from the provided XML string and appends it to the document.
        /// </summary>
        /// <param name="document">Target document.</param>
        /// <param name="xml">XML string representing a paragraph.</param>
        /// <returns>The inserted <see cref="WordParagraph"/>.</returns>
        public static WordParagraph AddParagraphFromXml(this WordDocument document, string xml) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            if (string.IsNullOrWhiteSpace(xml)) throw new ArgumentException("Value cannot be null or empty.", nameof(xml));
            var paragraph = new Paragraph(xml);
            var wordParagraph = new WordParagraph(document, paragraph);
            document.AddParagraph(wordParagraph);
            return wordParagraph;
        }
    }
}
