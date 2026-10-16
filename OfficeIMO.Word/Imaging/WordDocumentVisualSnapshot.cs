using OfficeIMO.Drawing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Shared page scene and text geometry for a Word document page.
    /// </summary>
    public sealed class WordDocumentVisualSnapshot {
        internal WordDocumentVisualSnapshot(
            OfficeDrawing drawing,
            int pageIndex,
            IReadOnlyList<OfficeImageExportDiagnostic> diagnostics,
            IReadOnlyList<OfficeIMO.Pdf.PdfTextSpan> textSpans) {
            Drawing = drawing;
            PageIndex = pageIndex;
            Diagnostics = diagnostics;
            TextSpans = textSpans;
        }

        /// <summary>Page drawing scene in points before export scaling.</summary>
        public OfficeDrawing Drawing { get; }

        /// <summary>Zero-based page index represented by this snapshot.</summary>
        public int PageIndex { get; }

        /// <summary>Snapshot diagnostics for Word content that could not be represented in the shared drawing scene.</summary>
        public IReadOnlyList<OfficeImageExportDiagnostic> Diagnostics { get; }

        /// <summary>Text and geometry from the shared page layout, including running headers and footers, in PDF user-space coordinates.</summary>
        public IReadOnlyList<OfficeIMO.Pdf.PdfTextSpan> TextSpans { get; }

        /// <summary>Snapshot width in points before export scaling.</summary>
        public double Width => Drawing.Width;

        /// <summary>Snapshot height in points before export scaling.</summary>
        public double Height => Drawing.Height;
    }
}
