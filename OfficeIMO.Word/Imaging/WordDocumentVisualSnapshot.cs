using OfficeIMO.Drawing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Format-neutral visual snapshot for a Word document page.
    /// </summary>
    public sealed class WordDocumentVisualSnapshot {
        internal WordDocumentVisualSnapshot(OfficeDrawing drawing, int pageIndex, IReadOnlyList<OfficeImageExportDiagnostic> diagnostics) {
            Drawing = drawing;
            PageIndex = pageIndex;
            Diagnostics = diagnostics;
        }

        /// <summary>Page drawing scene in points before export scaling.</summary>
        public OfficeDrawing Drawing { get; }

        /// <summary>Zero-based page index represented by this snapshot.</summary>
        public int PageIndex { get; }

        /// <summary>Snapshot diagnostics for Word content that could not be represented in the shared drawing scene.</summary>
        public IReadOnlyList<OfficeImageExportDiagnostic> Diagnostics { get; }

        /// <summary>Snapshot width in points before export scaling.</summary>
        public double Width => Drawing.Width;

        /// <summary>Snapshot height in points before export scaling.</summary>
        public double Height => Drawing.Height;
    }
}
