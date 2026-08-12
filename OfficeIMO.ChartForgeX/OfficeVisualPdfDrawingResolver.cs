using System;
using OfficeIMO.Drawing;

namespace OfficeIMO.ChartForgeX;

internal static class OfficeVisualPdfDrawingResolver {
    public static OfficeDrawing Resolve(OfficeVisualConversionResult conversion) {
        if (!RequiresRasterFallback(conversion.Drawing)) return conversion.Drawing;
        if (conversion.SvgPolicy == OfficeVisualSvgPolicy.RequireVector) {
            throw new NotSupportedException("The converted SVG contains a blend mode or soft mask that OfficeIMO.Pdf cannot preserve as vector content.");
        }

        byte[] png = OfficeDrawingRasterRenderer.ToPng(conversion.Drawing, scale: 96D / 72D);
        var drawing = new OfficeDrawing(conversion.WidthPoints, conversion.HeightPoints);
        drawing.AddImage(
            png,
            "image/png",
            new OfficeImageProjection(new OfficeImagePlacement(0D, 0D, conversion.WidthPoints, conversion.HeightPoints)),
            conversion.AlternativeText);
        return drawing;
    }

    private static bool RequiresRasterFallback(OfficeDrawing drawing) {
        foreach (OfficeDrawingElement element in drawing.Elements) {
            if (element is OfficeDrawingEffectGroup effectGroup) {
                if (effectGroup.BlendMode != OfficeBlendMode.Normal || effectGroup.SoftMask != null) return true;
                if (RequiresRasterFallback(effectGroup.Drawing)) return true;
            } else if (element is OfficeDrawingGroup group && RequiresRasterFallback(group.Drawing)) {
                return true;
            }
        }

        return false;
    }
}
