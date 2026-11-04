using OfficeIMO.Drawing;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// Format-neutral visual snapshot for a PowerPoint slide.
    /// </summary>
    public sealed class PowerPointSlideVisualInfo {
        internal PowerPointSlideVisualInfo(OfficeDrawing drawing, IReadOnlyList<OfficeImageExportDiagnostic> diagnostics) {
            Drawing = drawing;
            Diagnostics = diagnostics;
        }

        /// <summary>Slide drawing scene in points before export scaling.</summary>
        public OfficeDrawing Drawing { get; }

        /// <summary>Snapshot diagnostics for content that could not be represented in the shared drawing scene.</summary>
        public IReadOnlyList<OfficeImageExportDiagnostic> Diagnostics { get; }

        /// <summary>Native scene text geometry, measured by the shared Drawing layout.</summary>
        public IReadOnlyList<PowerPointTextSpanInfo> TextSpans => PowerPointSlideTextInspector.Read(Drawing);

        /// <summary>Snapshot width in points before export scaling.</summary>
        public double Width => Drawing.Width;

        /// <summary>Snapshot height in points before export scaling.</summary>
        public double Height => Drawing.Height;
    }
}
