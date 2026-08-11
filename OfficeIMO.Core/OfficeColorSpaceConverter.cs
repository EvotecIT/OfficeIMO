using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

/// <summary>Dependency-free conversion from common device and calibrated color spaces to sRGB.</summary>
public static class OfficeColorSpaceConverter {
    private const double D65X = 0.95047D;
    private const double D65Y = 1D;
    private const double D65Z = 1.08883D;

    /// <summary>Converts normalized subtractive device components to sRGB.</summary>
    public static OfficeColor FromCmyk(double cyan, double magenta, double yellow, double black) {
        double c = Clamp01(cyan);
        double m = Clamp01(magenta);
        double y = Clamp01(yellow);
        double k = Clamp01(black);
        return OfficeColor.FromRgb(ToByte((1D - c) * (1D - k)), ToByte((1D - m) * (1D - k)), ToByte((1D - y) * (1D - k)));
    }

    /// <summary>Converts CIE L*a*b* using the D50 reference white to sRGB.</summary>
    public static OfficeColor FromLab(double lightness, double a, double b) =>
        FromLab(lightness, a, b, 0.96422D, 1D, 0.82521D);

    /// <summary>Converts CIE L*a*b* using an explicit XYZ reference white to sRGB.</summary>
    public static OfficeColor FromLab(double lightness, double a, double b, double whiteX, double whiteY, double whiteZ) {
        ConvertLabToXyz(lightness, a, b, whiteX, whiteY, whiteZ, out double x, out double y, out double z);
        return FromXyz(x, y, z, whiteX, whiteY, whiteZ);
    }

    internal static void ConvertLabToXyz(
        double lightness,
        double a,
        double b,
        double whiteX,
        double whiteY,
        double whiteZ,
        out double x,
        out double y,
        out double z) {
        ValidateWhitePoint(whiteX, whiteY, whiteZ);
        double l = Clamp(lightness, 0D, 100D);
        double fy = (l + 16D) / 116D;
        double fx = fy + (Clamp(a, -128D, 127D) / 500D);
        double fz = fy - (Clamp(b, -128D, 127D) / 200D);
        x = whiteX * InverseLabPivot(fx);
        y = whiteY * InverseLabPivot(fy);
        z = whiteZ * InverseLabPivot(fz);
    }

    /// <summary>Converts a calibrated gray component and gamma using an explicit XYZ white point.</summary>
    public static OfficeColor FromCalibratedGray(double gray, double whiteX, double whiteY, double whiteZ, double gamma = 1D) {
        ValidateWhitePoint(whiteX, whiteY, whiteZ);
        ValidateGamma(gamma, nameof(gamma));
        double level = Math.Pow(Clamp01(gray), gamma);
        return FromXyz(whiteX * level, whiteY * level, whiteZ * level, whiteX, whiteY, whiteZ);
    }

    /// <summary>
    /// Converts calibrated RGB through per-channel gamma and a nine-value column-major XYZ matrix.
    /// Missing gamma and matrix values use PDF-compatible identity defaults.
    /// </summary>
    public static OfficeColor FromCalibratedRgb(
        double red,
        double green,
        double blue,
        double whiteX,
        double whiteY,
        double whiteZ,
        IReadOnlyList<double>? gamma = null,
        IReadOnlyList<double>? matrix = null) {
        ValidateWhitePoint(whiteX, whiteY, whiteZ);
        double gammaR = Component(gamma, 0, 1D);
        double gammaG = Component(gamma, 1, 1D);
        double gammaB = Component(gamma, 2, 1D);
        ValidateGamma(gammaR, nameof(gamma));
        ValidateGamma(gammaG, nameof(gamma));
        ValidateGamma(gammaB, nameof(gamma));
        double a = Math.Pow(Clamp01(red), gammaR);
        double b = Math.Pow(Clamp01(green), gammaG);
        double c = Math.Pow(Clamp01(blue), gammaB);
        double x = (Component(matrix, 0, 1D) * a) + (Component(matrix, 3, 0D) * b) + (Component(matrix, 6, 0D) * c);
        double y = (Component(matrix, 1, 0D) * a) + (Component(matrix, 4, 1D) * b) + (Component(matrix, 7, 0D) * c);
        double z = (Component(matrix, 2, 0D) * a) + (Component(matrix, 5, 0D) * b) + (Component(matrix, 8, 1D) * c);
        return FromXyz(x, y, z, whiteX, whiteY, whiteZ);
    }

    /// <summary>Converts XYZ values relative to an explicit source white point to sRGB.</summary>
    public static OfficeColor FromXyz(double x, double y, double z, double sourceWhiteX = D65X, double sourceWhiteY = D65Y, double sourceWhiteZ = D65Z) {
        ValidateWhitePoint(sourceWhiteX, sourceWhiteY, sourceWhiteZ);
        AdaptToD65(ref x, ref y, ref z, sourceWhiteX, sourceWhiteY, sourceWhiteZ);
        double linearRed = (3.2404542D * x) - (1.5371385D * y) - (0.4985314D * z);
        double linearGreen = (-0.969266D * x) + (1.8760108D * y) + (0.041556D * z);
        double linearBlue = (0.0556434D * x) - (0.2040259D * y) + (1.0572252D * z);
        return OfficeColor.FromRgb(ToSrgbByte(linearRed), ToSrgbByte(linearGreen), ToSrgbByte(linearBlue));
    }

    internal static void ConvertRgbToXyz(
        double red,
        double green,
        double blue,
        double targetWhiteX,
        double targetWhiteY,
        double targetWhiteZ,
        out double x,
        out double y,
        out double z) {
        ValidateWhitePoint(targetWhiteX, targetWhiteY, targetWhiteZ);
        double linearRed = FromSrgb(Clamp01(red));
        double linearGreen = FromSrgb(Clamp01(green));
        double linearBlue = FromSrgb(Clamp01(blue));
        x = (0.4124564D * linearRed) + (0.3575761D * linearGreen) + (0.1804375D * linearBlue);
        y = (0.2126729D * linearRed) + (0.7151522D * linearGreen) + (0.072175D * linearBlue);
        z = (0.0193339D * linearRed) + (0.119192D * linearGreen) + (0.9503041D * linearBlue);
        AdaptWhitePoint(ref x, ref y, ref z, D65X, D65Y, D65Z, targetWhiteX, targetWhiteY, targetWhiteZ);
    }

    private static void AdaptToD65(ref double x, ref double y, ref double z, double whiteX, double whiteY, double whiteZ) {
        AdaptWhitePoint(ref x, ref y, ref z, whiteX, whiteY, whiteZ, D65X, D65Y, D65Z);
    }

    private static void AdaptWhitePoint(
        ref double x,
        ref double y,
        ref double z,
        double sourceWhiteX,
        double sourceWhiteY,
        double sourceWhiteZ,
        double targetWhiteX,
        double targetWhiteY,
        double targetWhiteZ) {
        if (NearlyEqual(sourceWhiteX, targetWhiteX) && NearlyEqual(sourceWhiteY, targetWhiteY) && NearlyEqual(sourceWhiteZ, targetWhiteZ)) return;
        double sourceL = (0.8951D * sourceWhiteX) + (0.2664D * sourceWhiteY) - (0.1614D * sourceWhiteZ);
        double sourceM = (-0.7502D * sourceWhiteX) + (1.7135D * sourceWhiteY) + (0.0367D * sourceWhiteZ);
        double sourceS = (0.0389D * sourceWhiteX) - (0.0685D * sourceWhiteY) + (1.0296D * sourceWhiteZ);
        double targetL = (0.8951D * targetWhiteX) + (0.2664D * targetWhiteY) - (0.1614D * targetWhiteZ);
        double targetM = (-0.7502D * targetWhiteX) + (1.7135D * targetWhiteY) + (0.0367D * targetWhiteZ);
        double targetS = (0.0389D * targetWhiteX) - (0.0685D * targetWhiteY) + (1.0296D * targetWhiteZ);
        double l = ((0.8951D * x) + (0.2664D * y) - (0.1614D * z)) * (targetL / sourceL);
        double m = ((-0.7502D * x) + (1.7135D * y) + (0.0367D * z)) * (targetM / sourceM);
        double s = ((0.0389D * x) - (0.0685D * y) + (1.0296D * z)) * (targetS / sourceS);
        x = (0.9869929D * l) - (0.1470543D * m) + (0.1599627D * s);
        y = (0.4323053D * l) + (0.5183603D * m) + (0.0492912D * s);
        z = (-0.0085287D * l) + (0.0400428D * m) + (0.9684867D * s);
    }

    private static double InverseLabPivot(double value) {
        double cube = value * value * value;
        return cube > 216D / 24389D ? cube : (116D * value - 16D) / 903.3D;
    }
    private static byte ToSrgbByte(double linear) {
        double value = linear <= 0.0031308D ? 12.92D * linear : (1.055D * Math.Pow(Math.Max(0D, linear), 1D / 2.4D)) - 0.055D;
        return ToByte(value);
    }
    private static double FromSrgb(double value) => value <= 0.04045D
        ? value / 12.92D
        : Math.Pow((value + 0.055D) / 1.055D, 2.4D);
    private static byte ToByte(double value) => (byte)Math.Round(Clamp01(value) * 255D);
    private static double Component(IReadOnlyList<double>? values, int index, double fallback) => values != null && index < values.Count ? values[index] : fallback;
    private static double Clamp01(double value) => Clamp(value, 0D, 1D);
    private static double Clamp(double value, double minimum, double maximum) => value < minimum ? minimum : value > maximum ? maximum : value;
    private static bool NearlyEqual(double left, double right) => Math.Abs(left - right) <= 0.00001D;
    private static void ValidateWhitePoint(double x, double y, double z) {
        if (!IsFinite(x) || !IsFinite(y) || !IsFinite(z) || x <= 0D || y <= 0D || z <= 0D) throw new ArgumentOutOfRangeException(nameof(x), "XYZ white-point components must be finite and positive.");
    }
    private static void ValidateGamma(double value, string name) {
        if (!IsFinite(value) || value <= 0D) throw new ArgumentOutOfRangeException(name, "Color-space gamma must be finite and positive.");
    }
    private static bool IsFinite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);
}
