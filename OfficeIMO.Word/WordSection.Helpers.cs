using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Convenience helpers for working with section-scoped headers/footers.
    /// </summary>
    public partial class WordSection {
        /// <summary>
        /// Returns the section header of the requested type (Default/Even/First).
        /// </summary>
        public WordHeader? GetHeader(WordHeaderFooterType type) {
            if (type == WordHeaderFooterType.First) return this.Header.First;
            if (type == WordHeaderFooterType.Even) return this.Header.Even;
            return this.Header.Default;
        }
        /// <summary>
        /// Returns the default section header.
        /// </summary>
        public WordHeader? GetHeader() => GetHeader(WordHeaderFooterType.Default);

        /// <summary>
        /// Returns the section footer of the requested type (Default/Even/First).
        /// </summary>
        public WordFooter? GetFooter(WordHeaderFooterType type) {
            if (type == WordHeaderFooterType.First) return this.Footer.First;
            if (type == WordHeaderFooterType.Even) return this.Footer.Even;
            return this.Footer.Default;
        }
        /// <summary>
        /// Returns the default section footer.
        /// </summary>
        public WordFooter? GetFooter() => GetFooter(WordHeaderFooterType.Default);

        /// <summary>
        /// Adds a paragraph to the section header of the requested type.
        /// </summary>
        public WordParagraph AddHeaderParagraph(string text, WordHeaderFooterType type, bool removeExistingParagraphs = false) {
            var header = GetHeader(type) ?? throw new InvalidOperationException("Header not available for this section.");
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
            var footer = GetFooter(type) ?? throw new InvalidOperationException("Footer not available for this section.");
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
