using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Options for headless SVG export from an in-memory Visio page.
    /// </summary>
    public sealed class VisioSvgSaveOptions {
        /// <summary>
        /// Zero-based page index used when exporting a document. Defaults to the first page.
        /// </summary>
        public int PageIndex { get; set; }

        /// <summary>
        /// Number of SVG units used per Visio inch. Defaults to 96 so the SVG maps naturally to browser pixels.
        /// </summary>
        public double PixelsPerInch { get; set; } = 96D;

        /// <summary>
        /// Optional page background. Defaults to white; set to <c>null</c> for transparent output.
        /// </summary>
        public Color? BackgroundColor { get; set; } = Color.White;

        /// <summary>
        /// Gets or sets whether shape text is emitted.
        /// </summary>
        public bool RenderText { get; set; } = true;

        /// <summary>
        /// Gets or sets whether connector labels are emitted.
        /// </summary>
        public bool RenderConnectorLabels { get; set; } = true;

        /// <summary>
        /// Gets or sets whether an XML declaration should be included.
        /// </summary>
        public bool IncludeXmlDeclaration { get; set; }
    }
}
