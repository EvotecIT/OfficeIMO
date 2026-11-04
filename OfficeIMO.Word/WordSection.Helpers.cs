using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    // Convenience helpers for working with section-scoped headers/footers.
    public partial class WordSection {
        /// <summary>
        /// Adds a paragraph to the section header of the requested type.
        /// </summary>
        public WordParagraph AddHeaderParagraph(string text, WordHeaderFooterType type, bool removeExistingParagraphs = false) {
            var header = PageSetup.GetHeader(type) ?? throw new InvalidOperationException("Header not available for this section.");
            if (removeExistingParagraphs) {
                // Clear existing header paragraphs
                foreach (var p in header.Paragraphs.ToList()) p.Remove();
            }
            return string.IsNullOrEmpty(text) ? header.AddParagraph("") : header.AddParagraph(text);
        }
        /// <summary>
        /// Adds a paragraph to the default section header.
        /// </summary>
        /// <param name="text">Paragraph text.</param>
        /// <param name="removeExistingParagraphs">True to clear existing paragraphs.</param>
        public WordParagraph AddHeaderParagraph(string text = "", bool removeExistingParagraphs = false) =>
            AddHeaderParagraph(text, WordHeaderFooterType.Default, removeExistingParagraphs);

        /// <summary>
        /// Adds a paragraph to the section footer of the requested type.
        /// </summary>
        public WordParagraph AddFooterParagraph(string text, WordHeaderFooterType type, bool removeExistingParagraphs = false) {
            var footer = PageSetup.GetFooter(type) ?? throw new InvalidOperationException("Footer not available for this section.");
            if (removeExistingParagraphs) {
                foreach (var p in footer.Paragraphs.ToList()) p.Remove();
            }
            return string.IsNullOrEmpty(text) ? footer.AddParagraph("") : footer.AddParagraph(text);
        }
        /// <summary>
        /// Adds a paragraph to the default section footer.
        /// </summary>
        /// <param name="text">Paragraph text.</param>
        /// <param name="removeExistingParagraphs">True to clear existing paragraphs.</param>
        public WordParagraph AddFooterParagraph(string text = "", bool removeExistingParagraphs = false) =>
            AddFooterParagraph(text, WordHeaderFooterType.Default, removeExistingParagraphs);
    }
}
