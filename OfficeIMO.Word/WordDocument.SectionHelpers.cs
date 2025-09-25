using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Convenience helpers for addressing section-scoped headers/footers from the document root.
    /// </summary>
    public partial class WordDocument {
        /// <summary>
        /// Returns the header for a specific section index (use -1 to target the last section).
        /// </summary>
        public WordHeader? GetHeaderForSection(int sectionIndex, HeaderFooterValues type) {
            if (this.Sections.Count == 0) throw new InvalidOperationException("Document has no sections.");
            if (sectionIndex < 0) sectionIndex = this.Sections.Count - 1;
            if (sectionIndex >= this.Sections.Count) throw new ArgumentOutOfRangeException(nameof(sectionIndex));
            var headers = this.Sections[sectionIndex].Header;
            if (type == HeaderFooterValues.First) return headers.First;
            if (type == HeaderFooterValues.Even) return headers.Even;
            return headers.Default;
        }
        /// <summary>
        /// Returns the default header for a specific section index.
        /// </summary>
        /// <param name="sectionIndex">Index of the section or -1 for the last section.</param>
        public WordHeader? GetHeaderForSection(int sectionIndex = -1) => GetHeaderForSection(sectionIndex, HeaderFooterValues.Default);

        /// <summary>
        /// Returns the footer for a specific section index (use -1 to target the last section).
        /// </summary>
        public WordFooter? GetFooterForSection(int sectionIndex, HeaderFooterValues type) {
            if (this.Sections.Count == 0) throw new InvalidOperationException("Document has no sections.");
            if (sectionIndex < 0) sectionIndex = this.Sections.Count - 1;
            if (sectionIndex >= this.Sections.Count) throw new ArgumentOutOfRangeException(nameof(sectionIndex));
            var footers = this.Sections[sectionIndex].Footer;
            if (type == HeaderFooterValues.First) return footers.First;
            if (type == HeaderFooterValues.Even) return footers.Even;
            return footers.Default;
        }
        /// <summary>
        /// Returns the default footer for a specific section index.
        /// </summary>
        /// <param name="sectionIndex">Index of the section or -1 for the last section.</param>
        public WordFooter? GetFooterForSection(int sectionIndex = -1) => GetFooterForSection(sectionIndex, HeaderFooterValues.Default);
    }
}
