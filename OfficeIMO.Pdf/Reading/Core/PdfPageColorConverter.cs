using System;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static class PdfPageColorConverter {
    public static OfficeColor FromLab(double lightness, double a, double b) {
        double l = Clamp(lightness, 0D, 100D);
        double fy = (l + 16D) / 116D;
        double fx = fy + (Clamp(a, -128D, 127D) / 500D);
        double fz = fy - (Clamp(b, -128D, 127D) / 200D);

        double x50 = 0.96422D * InverseLabPivot(fx);
        double y50 = InverseLabPivot(fy);
        double z50 = 0.82521D * InverseLabPivot(fz);

        double x = (0.9555766D * x50) - (0.0230393D * y50) + (0.0631636D * z50);
        double y = (-0.0282895D * x50) + (1.0099416D * y50) + (0.0210077D * z50);
        double z = (0.0122982D * x50) - (0.020483D * y50) + (1.3299098D * z50);

        double red = (3.2404542D * x) - (1.5371385D * y) - (0.4985314D * z);
        double green = (-0.969266D * x) + (1.8760108D * y) + (0.041556D * z);
        double blue = (0.0556434D * x) - (0.2040259D * y) + (1.0572252D * z);
        return OfficeColor.FromRgb(ToSrgbByte(red), ToSrgbByte(green), ToSrgbByte(blue));
    }

    private static double InverseLabPivot(double value) {
        double cube = value * value * value;
        return cube > 216D / 24389D ? cube : (116D * value - 16D) / 903.3D;
    }

    private static byte ToSrgbByte(double linear) {
        double value = linear <= 0.0031308D
            ? 12.92D * linear
            : (1.055D * Math.Pow(Math.Max(0D, linear), 1D / 2.4D)) - 0.055D;
        return (byte)Math.Round(Clamp(value, 0D, 1D) * 255D);
    }

    private static double Clamp(double value, double minimum, double maximum) =>
        value < minimum ? minimum : value > maximum ? maximum : value;
}
