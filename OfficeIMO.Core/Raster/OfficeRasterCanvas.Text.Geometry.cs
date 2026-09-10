using System;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeRasterCanvas {
    private static double ResolveRasterTextSize(double fontSize, double height, bool positioned) =>
        positioned ? Math.Max(0.1D, fontSize) : Math.Max(6D, Math.Min(fontSize, height - 2D));

    private static double ResolveRasterTextTop(IOfficeFontProgram font, double size, double height, double? baselineFontSize) {
        double sourceSize = baselineFontSize ?? size;
        double top = Math.Max(1D, (height - font.LineHeight(sourceSize)) / 2D);
        if (Math.Abs(sourceSize - size) < 0.000001D) return top;
        // A smaller script glyph keeps the source line's baseline, then applies its own offset.
        return top + ResolveRasterBaseline(font, sourceSize) - ResolveRasterBaseline(font, size);
    }

    private static double ResolveRasterBaseline(IOfficeFontProgram font, double size) {
        double lineHeight = font.LineHeight(size);
        double baseline = font is IOfficeFontBaselineMetrics metrics ? metrics.BaselineOffset(size) : lineHeight * 0.8D;
        return double.IsNaN(baseline) || double.IsInfinity(baseline)
            ? lineHeight * 0.8D : Math.Max(0D, Math.Min(lineHeight, baseline));
    }
}
