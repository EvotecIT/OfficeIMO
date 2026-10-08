using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

internal static class PdfAffineRadialAssertions {
    internal static void RetainsVectorRadialPaint(byte[] pdf, int minimumRadialShapes = 1) {
        var result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic =>
            diagnostic.Code == PdfRenderCapabilities.UnsupportedShadingId || diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        var drawing = PdfPageImageRenderer.RenderPage(pdf);
        Assert.True(RadialShapes(drawing).Count() >= minimumRadialShapes,
            "Affine radial paint must remain vector content rather than a font-substitution or empty fallback.");
    }

    private static IEnumerable<OfficeDrawingShape> RadialShapes(OfficeDrawing drawing) {
        foreach (var element in drawing.Elements) {
            if (element is OfficeDrawingShape shape && (shape.Shape.FillRadialGradient != null || shape.Shape.StrokeRadialGradient != null)) yield return shape;
            OfficeDrawing? nested = element switch {
                OfficeDrawingGroup group => group.Drawing,
                OfficeDrawingEffectGroup effect => effect.Drawing,
                OfficeDrawingTilingPattern pattern => pattern.Tile,
                _ => null
            };
            if (nested != null) foreach (var child in RadialShapes(nested)) yield return child;
        }
    }
}
