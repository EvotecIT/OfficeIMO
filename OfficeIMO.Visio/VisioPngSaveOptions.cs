using System.Collections.Generic;
using OfficeIMO.Drawing;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Options for native dependency-free PNG export from an in-memory Visio page.
    /// </summary>
    public sealed class VisioPngSaveOptions {
        /// <summary>
        /// Zero-based page index used when exporting a document. Defaults to the first page.
        /// </summary>
        public int PageIndex { get; set; }

        /// <summary>
        /// Number of output pixels used per Visio inch. Defaults to 96.
        /// </summary>
        public double PixelsPerInch { get; set; } = 96D;

        /// <summary>
        /// Optional page background. Defaults to white; set to <c>null</c> for transparent output.
        /// </summary>
        public Color? BackgroundColor { get; set; } = Color.White;

        /// <summary>
        /// Gets or sets whether shape text is rendered.
        /// </summary>
        public bool RenderText { get; set; } = true;

        /// <summary>
        /// Optional TrueType/OpenType font file used for native PNG text outlines. When unset or unreadable, the renderer uses the built-in managed fallback discovery path.
        /// </summary>
        public string? FontFilePath { get; set; }

        /// <summary>
        /// Optional font face name used when selecting a face from a TrueType/OpenType collection.
        /// </summary>
        public string? FontFaceName { get; set; }

        /// <summary>
        /// Optional zero-based face index used when selecting a face from a TrueType/OpenType collection.
        /// </summary>
        public int? FontCollectionIndex { get; set; }

        /// <summary>Caller-supplied deterministic TrueType faces used before platform fallback.</summary>
        public OfficeFontFaceCollection Fonts { get; set; } = new OfficeFontFaceCollection();

        /// <summary>Cancellation observed between shapes and connectors.</summary>
        public System.Threading.CancellationToken CancellationToken { get; set; }

        /// <summary>
        /// Gets or sets whether built-in OfficeIMO stencil metadata is projected as dependency-free vector artwork.
        /// </summary>
        public bool RenderStencilArtwork { get; set; } = true;

        /// <summary>
        /// Gets or sets whether connector labels are rendered.
        /// </summary>
        public bool RenderConnectorLabels { get; set; } = true;

        /// <summary>
        /// Gets or sets whether connector labels are nudged at render time to avoid page edges, unrelated shapes, and earlier labels.
        /// </summary>
        public bool ResolveConnectorLabelOverlaps { get; set; } = true;

        /// <summary>
        /// Supersampling factor used for smoother native raster output. Defaults to 3.
        /// </summary>
        public int Supersampling { get; set; } = 3;

        internal IOfficeRasterImageCodec? ImageCodec { get; set; }

        internal ICollection<OfficeImageExportDiagnostic>? ImageDiagnostics { get; set; }

        internal string? ImageDiagnosticSource { get; set; }
    }
}
