using System;

namespace OfficeIMO.Drawing;

internal static class OfficeGradientColors {
    internal static OfficeColor Interpolate(OfficeColor start, OfficeColor end, double ratio, OfficeGradientColorInterpolation mode) {
        ratio = Math.Max(0, Math.Min(1, ratio));
        byte alpha = (byte)Math.Round(start.A + (end.A - start.A) * ratio);
        if (mode == OfficeGradientColorInterpolation.LinearRgb) {
            var rgb = OfficeColorSpaceConverter.FromLinearSrgb(
                LinearChannel(start.R, end.R, ratio), LinearChannel(start.G, end.G, ratio), LinearChannel(start.B, end.B, ratio));
            return OfficeColor.FromRgba(rgb.R, rgb.G, rgb.B, alpha);
        }
        return OfficeColor.FromRgba((byte)Math.Round(start.R + (end.R - start.R) * ratio),
            (byte)Math.Round(start.G + (end.G - start.G) * ratio), (byte)Math.Round(start.B + (end.B - start.B) * ratio), alpha);
    }

    internal static OfficeColor Average(System.Collections.Generic.IReadOnlyList<OfficeGradientStop> stops, OfficeGradientColorInterpolation mode) {
        double red = 0, green = 0, blue = 0, alpha = 0;
        double Channel(byte value) => mode == OfficeGradientColorInterpolation.LinearRgb ? OfficeColorSpaceConverter.FromSrgb(value / 255D) : value / 255D;
        for (int i = 1; i < stops.Count; i++) {
            var a = stops[i - 1]; var b = stops[i]; double weight = (b.Offset - a.Offset) / 2D;
            red += weight * (Channel(a.Color.R) + Channel(b.Color.R));
            green += weight * (Channel(a.Color.G) + Channel(b.Color.G));
            blue += weight * (Channel(a.Color.B) + Channel(b.Color.B));
            alpha += weight * (a.Color.A + b.Color.A);
        }
        byte Byte(double value) => (byte)Math.Round(Math.Max(0, Math.Min(255, value)));
        var rgb = mode == OfficeGradientColorInterpolation.LinearRgb ? OfficeColorSpaceConverter.FromLinearSrgb(red, green, blue)
            : OfficeColor.FromRgb(Byte(red * 255), Byte(green * 255), Byte(blue * 255));
        return OfficeColor.FromRgba(rgb.R, rgb.G, rgb.B, Byte(alpha));
    }

    private static double LinearChannel(byte first, byte second, double ratio) {
        double a = OfficeColorSpaceConverter.FromSrgb(first / 255D), b = OfficeColorSpaceConverter.FromSrgb(second / 255D);
        return a + (b - a) * ratio;
    }
}
