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

    private static double LinearChannel(byte first, byte second, double ratio) {
        double a = OfficeColorSpaceConverter.FromSrgb(first / 255D), b = OfficeColorSpaceConverter.FromSrgb(second / 255D);
        return a + (b - a) * ratio;
    }
}
