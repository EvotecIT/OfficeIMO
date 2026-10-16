using OfficeIMO.Drawing;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// Format-neutral visual snapshot for a PowerPoint slide.
    /// </summary>
    public sealed class PowerPointSlideVisualSnapshot {
        internal PowerPointSlideVisualSnapshot(OfficeDrawing drawing, IReadOnlyList<OfficeImageExportDiagnostic> diagnostics, IReadOnlyList<OfficeIMO.Pdf.PdfTextSpan> textSpans) {
            Drawing = drawing;
            Diagnostics = diagnostics;
            TextSpans = textSpans;
        }

        /// <summary>Slide drawing scene in points before export scaling.</summary>
        public OfficeDrawing Drawing { get; }

        /// <summary>Snapshot diagnostics for content that could not be represented in the shared drawing scene.</summary>
        public IReadOnlyList<OfficeImageExportDiagnostic> Diagnostics { get; }

        /// <summary>Text geometry from the canonical PDF slide layout, in PDF user-space coordinates.</summary>
        public IReadOnlyList<OfficeIMO.Pdf.PdfTextSpan> TextSpans { get; }

        /// <summary>Snapshot width in points before export scaling.</summary>
        public double Width => Drawing.Width;

        /// <summary>Snapshot height in points before export scaling.</summary>
        public double Height => Drawing.Height;
    }
}
