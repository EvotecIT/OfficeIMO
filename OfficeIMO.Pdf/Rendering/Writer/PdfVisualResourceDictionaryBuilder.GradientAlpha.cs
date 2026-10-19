using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfVisualResourceDictionaryBuilder {
    private static bool HasVaryingGradientAlpha(IReadOnlyList<OfficeGradientStop> stops) {
        for (int i = 1; i < stops.Count; i++)
            if (stops[i - 1].Color.A != stops[i].Color.A) return true;
        return false;
    }

    // Interpolate in the selected color space, premultiplying before interpolation
    // for canvas paint. SVG/XPS imports retain their independent color/alpha intent.
    private static double[] GradientComponents(OfficeColor start, OfficeColor end, double position,
        PdfPrintColorTransform? printColorTransform, OfficeGradientColorInterpolation interpolation,
        bool interpolateAlphaSeparately) {
        double startAlpha = start.A / 255D, endAlpha = end.A / 255D;
        double alpha = startAlpha + (endAlpha - startAlpha) * position;
        double Channel(byte first, byte second) {
            double a = first / 255D, b = second / 255D;
            if (interpolation == OfficeGradientColorInterpolation.LinearRgb) {
                a = OfficeColorSpaceConverter.FromSrgb(a);
                b = OfficeColorSpaceConverter.FromSrgb(b);
            }
            if (interpolateAlphaSeparately) return a + (b - a) * position;
            // Use the visible endpoint's limiting color when alpha is zero. This
            // avoids adding an invisible RGB fringe next to transparent stops.
            if (alpha <= 0D) return startAlpha > 0D ? a : endAlpha > 0D ? b : 0D;
            return (a * startAlpha * (1D - position) + b * endAlpha * position) / alpha;
        }
        double red = Channel(start.R, end.R), green = Channel(start.G, end.G), blue = Channel(start.B, end.B);
        if (printColorTransform == null) return new[] { red, green, blue };
        OfficeColor color = interpolation == OfficeGradientColorInterpolation.LinearRgb
            ? OfficeColorSpaceConverter.FromLinearSrgb(red, green, blue)
            : OfficeColor.FromRgb(ToByte(red), ToByte(green), ToByte(blue));
        var components = new double[4];
        printColorTransform.Convert(color, components);
        return components;
    }

    private static byte ToByte(double value) => (byte)Math.Round(Math.Max(0D, Math.Min(1D, value)) * 255D, MidpointRounding.AwayFromZero);
}
