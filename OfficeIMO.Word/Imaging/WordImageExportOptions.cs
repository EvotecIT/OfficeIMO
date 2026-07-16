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
        /// Gets or sets the page index to render. The current renderer supports the first page only.
        /// </summary>
        public int PageIndex { get; set; }

        internal WordImageExportOptions Clone() => new WordImageExportOptions {
            Scale = Scale,
            BackgroundColor = BackgroundColor,
            RasterEncoding = RasterEncoding?.Clone() ?? new OfficeRasterEncodingOptions(),
            IncludeDocumentContent = IncludeDocumentContent,
            PageIndex = PageIndex
        };
    }
}
