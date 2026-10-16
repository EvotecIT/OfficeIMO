using System;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        /// <summary>Paints a bounded vector halo before the source shadow and foreground.</summary>
        private void DrawShapeGlowAt(OfficeShape shape, double xShape, double bottomY) {
            OfficeGlow? glow = shape.Glow;
            if (glow == null || glow.Radius <= 0D || glow.Opacity <= 0D || glow.Color.A == 0) return;

            // Reuse the shared contour-layer planner without a filled shadow core.
            // Folding color alpha into the planned opacity keeps the total coverage bounded.
            var layers = OfficeShadowLayerPlanner.Create(
                glow.Opacity * glow.Color.A / 255D,
                glow.Radius,
                Math.Max(0D, shape.StrokeWidth),
                hasFill: false,
                hasStroke: true,
                canExpand: false);
            OfficeShape halo = shape.Clone();
            halo.Glow = null;
            halo.Shadow = null;
            halo.FillColor = null;
            halo.FillGradient = null;
            halo.FillRadialGradient = null;
            halo.FillOpacity = 1D;
            halo.StrokeColor = OfficeColor.FromRgb(glow.Color.R, glow.Color.G, glow.Color.B);
            halo.StrokeGradient = null;
            halo.StrokeRadialGradient = null;
            halo.StrokeDashStyle = OfficeStrokeDashStyle.Solid;
            halo.SetStrokeDashArray(null);
            halo.StrokeStartMarker = null;
            halo.StrokeEndMarker = null;
            for (int index = 0; index < layers.Count; index++) {
                cancellationToken.ThrowIfCancellationRequested();
                halo.StrokeWidth = layers[index].StrokeWidth;
                halo.StrokeOpacity = layers[index].Opacity;
                // The normal shape writer preserves local transforms, clipping and paint state.
                DrawShapeGeometryAt(halo, xShape, bottomY);
            }
        }
    }
}
