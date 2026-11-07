using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        /// <summary>Paints the block's existing detached scene through the same scale group without cloning it into a temporary drawing.</summary>
        private void DrawScaledCanvasDrawing(OfficeDrawing drawing, double scaleX, double scaleY,
            double originX, double originTopY) {
            var metrics = new DrawingTextMetricsState(currentOpts, cancellationToken);
            cancellationToken.ThrowIfCancellationRequested();
            OfficeTransform transform = ToTopLeftPageTransform(OfficeTransform.Scale(scaleX, scaleY), originX, originTopY);
            RenderEffectGroup(transform, 1D,
                () => DrawDrawingElements(drawing, originX, originTopY, metrics));
        }
    }
}
