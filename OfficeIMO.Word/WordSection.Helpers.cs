using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Convenience helpers for working with section-scoped headers/footers.
    /// </summary>
    public partial class WordSection {
        /// <summary>
        /// Returns the section header of the requested type (Default/Even/First).
        /// </summary>
        public WordHeader GetHeader(HeaderFooterValues type) {
            if (type == HeaderFooterValues.First) return this.Header.First;
            if (type == HeaderFooterValues.Even) return this.Header.Even;
            return this.Header.Default;
        }
        public WordHeader GetHeader() => GetHeader(HeaderFooterValues.Default);

        /// <summary>
        /// Returns the section footer of the requested type (Default/Even/First).
        /// </summary>
        public WordFooter GetFooter(HeaderFooterValues type) {
            if (type == HeaderFooterValues.First) return this.Footer.First;
            if (type == HeaderFooterValues.Even) return this.Footer.Even;
            return this.Footer.Default;
        }
        public WordFooter GetFooter() => GetFooter(HeaderFooterValues.Default);

        /// <summary>
        /// Adds a paragraph to the section header of the requested type.
        /// </summary>
        public WordParagraph AddHeaderParagraph(string text, HeaderFooterValues type, bool removeExistingParagraphs = false) {
            var header = GetHeader(type);
            if (removeExistingParagraphs) {
                // Clear existing header paragraphs
                foreach (var p in header.Paragraphs.ToList()) p.Remove();
            }
            return string.IsNullOrEmpty(text) ? header.AddParagraph("") : header.AddParagraph(text);
        }
        public WordParagraph AddHeaderParagraph(string text = "", bool removeExistingParagraphs = false) =>
            AddHeaderParagraph(text, HeaderFooterValues.Default, removeExistingParagraphs);

        /// <summary>
        /// Adds a paragraph to the section footer of the requested type.
        /// </summary>
        public WordParagraph AddFooterParagraph(string text, HeaderFooterValues type, bool removeExistingParagraphs = false) {
            var footer = GetFooter(type);
            if (removeExistingParagraphs) {
                foreach (var p in footer.Paragraphs.ToList()) p.Remove();
            }
            return string.IsNullOrEmpty(text) ? footer.AddParagraph("") : footer.AddParagraph(text);
        }
        public WordParagraph AddFooterParagraph(string text = "", bool removeExistingParagraphs = false) =>
            AddFooterParagraph(text, HeaderFooterValues.Default, removeExistingParagraphs);
    }
}
