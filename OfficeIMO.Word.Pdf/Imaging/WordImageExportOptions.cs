using OfficeIMO.Drawing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Options controlling dependency-free Word document image export.
    /// </summary>
    public sealed class WordImageExportOptions : OfficeImageExportOptions {
        /// <inheritdoc />
        public override double LogicalUnitsPerInch => 72D;

        /// <summary>
        /// Gets or sets the shared Word page-layout settings used by both PDF and image export.
        /// Image-export fonts and shaping settings overlay these settings before pagination.
        /// </summary>
        public Pdf.WordPdfExportOptions LayoutOptions { get; set; } = new Pdf.WordPdfExportOptions();

        /// <summary>
        /// Gets or sets the zero-based page index to render or the first page used by batch export.
        /// </summary>
        public int PageIndex { get; set; }

        /// <summary>
        /// Gets or sets the maximum number of pages returned by batch export. A null value exports from
        /// <see cref="PageIndex"/> through the laid-out end of the document. Single-page export ignores this value.
        /// </summary>
        public int? PageCount { get; set; }

        /// <summary>Creates an independent options snapshot.</summary>
        public WordImageExportOptions Clone() {
            if (LayoutOptions == null) throw new ArgumentNullException(nameof(LayoutOptions));
            WordImageExportOptions clone = CopyImageExportOptionsTo(new WordImageExportOptions());
            clone.LayoutOptions = LayoutOptions.Clone();
            clone.PageIndex = PageIndex;
            clone.PageCount = PageCount;
            return clone;
        }

        internal void Validate() {
            ValidateImageExportOptions();
            if (LayoutOptions == null) throw new ArgumentNullException(nameof(LayoutOptions));
            if (PageIndex < 0) throw new ArgumentOutOfRangeException(nameof(PageIndex));
            if (PageCount.HasValue && PageCount.Value < 1) throw new ArgumentOutOfRangeException(nameof(PageCount));
        }

        internal OfficeIMO.Pdf.PdfImageExportOptions ToPdfImageOptions() =>
            CopyImageExportOptionsTo(new OfficeIMO.Pdf.PdfImageExportOptions());
    }
}
