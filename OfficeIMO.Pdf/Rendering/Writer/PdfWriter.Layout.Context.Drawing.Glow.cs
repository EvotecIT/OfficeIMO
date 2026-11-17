using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        /// <summary>Paints the blurred glow halo before the source shadow and foreground.</summary>
        private void DrawShapeGlowAt(OfficeShape shape, double xShape, double bottomY) {
            OfficeBlurredShapeEffect? glow = OfficeBlurredShapeEffect.CreateGlow(shape);
            if (glow != null) DrawBlurredShapeEffectAt(glow, xShape, bottomY, shape.Height);
        }
    }
}
