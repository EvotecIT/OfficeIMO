namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        /// <summary>Reports page clipping after the actual paragraph and margin transforms have been resolved.</summary>
        private void ReportAnchoredImagePageClipping(PdfCanvasImageItem item, double bottom) {
            (double left, double lower, double right, double upper) = _canvasEffectToPage.TransformRectangleBounds(
                item.X, bottom, item.Block.Width, item.Block.Height);
            if (left >= -0.001D && lower >= -0.001D &&
                right <= currentOpts.PageWidth.ToPoints() + 0.001D && upper <= currentOpts.PageHeight.ToPoints() + 0.001D) return;
            currentOpts.AddLayoutDiagnostic("anchored-image-page-bounds-clipped", "anchored image",
                "The anchored image extends beyond the physical page and is clipped by its boundary.",
                PdfLayoutDiagnosticKind.ClippedContent, x: left, y: lower,
                width: right - left, height: upper - lower);
        }
    }
}
