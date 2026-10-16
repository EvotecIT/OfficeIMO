using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        /// <summary>Paints a bounded vector halo before the source shadow and foreground.</summary>
        private void DrawShapeGlowAt(OfficeShape shape, double xShape, double bottomY) {
            var layers = OfficeGlowLayerPlanner.CreateShapes(shape);
            for (int index = 0; index < layers.Count; index++) {
                cancellationToken.ThrowIfCancellationRequested();
                // The normal shape writer preserves local transforms, clipping and paint state.
                DrawShapeGeometryAt(layers[index], xShape, bottomY);
            }
        }
    }
}
