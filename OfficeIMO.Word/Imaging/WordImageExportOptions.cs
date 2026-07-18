using OfficeIMO.Drawing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Options controlling dependency-free Word document image export.
    /// </summary>
    public sealed class WordImageExportOptions : OfficeImageExportOptions {
        /// <summary>
        /// Gets or sets a value indicating whether document body content should be rendered when supported.
        /// </summary>
        public bool IncludeDocumentContent { get; set; } = true;

        /// <summary>
        /// Gets or sets the zero-based page index to render or the first page used by batch export.
        /// </summary>
        public int PageIndex { get; set; }

        /// <summary>
        /// Gets or sets the maximum number of pages returned by batch export. A null value exports from
        /// <see cref="PageIndex"/> through the estimated end of the document. Single-page export ignores this value.
        /// </summary>
        public int? PageCount { get; set; }

        internal WordImageExportOptions Clone() {
            WordImageExportOptions clone = CopyImageExportOptionsTo(new WordImageExportOptions());
            clone.IncludeDocumentContent = IncludeDocumentContent;
            clone.PageIndex = PageIndex;
            clone.PageCount = PageCount;
            return clone;
        }

        internal void Validate() {
            ValidateImageExportOptions();
            if (PageIndex < 0) throw new ArgumentOutOfRangeException(nameof(PageIndex));
            if (PageCount.HasValue && PageCount.Value < 1) throw new ArgumentOutOfRangeException(nameof(PageCount));
        }
    }
}
