using System;

namespace OfficeIMO.Drawing;

/// <summary>WCAG-compatible relative luminance and contrast calculations for shared OfficeIMO colors.</summary>
public static class OfficeColorContrast {
    /// <summary>Returns the relative luminance of an sRGB color in the inclusive range 0 to 1.</summary>
    public static double RelativeLuminance(OfficeColor color) {
        double red = Linearize(color.R / 255D);
        double green = Linearize(color.G / 255D);
        double blue = Linearize(color.B / 255D);
        return 0.2126D * red + 0.7152D * green + 0.0722D * blue;
    }

    /// <summary>Returns the contrast ratio between two colors in the inclusive range 1 to 21.</summary>
    public static double ContrastRatio(OfficeColor first, OfficeColor second) {
        double firstLuminance = RelativeLuminance(first);
        double secondLuminance = RelativeLuminance(second);
        double lighter = Math.Max(firstLuminance, secondLuminance);
        double darker = Math.Min(firstLuminance, secondLuminance);
        return (lighter + 0.05D) / (darker + 0.05D);
    }

    /// <summary>Returns whether two colors meet the supplied minimum contrast ratio.</summary>
    public static bool Meets(OfficeColor foreground, OfficeColor background, double minimumRatio) {
        if (double.IsNaN(minimumRatio) || double.IsInfinity(minimumRatio) || minimumRatio < 1D || minimumRatio > 21D) {
            throw new ArgumentOutOfRangeException(nameof(minimumRatio), "Contrast ratio must be between 1 and 21.");
        }
        return ContrastRatio(foreground, background) + 0.000001D >= minimumRatio;
    }

    private static double Linearize(double channel) =>
        channel <= 0.04045D
            ? channel / 12.92D
            : Math.Pow((channel + 0.055D) / 1.055D, 2.4D);
}
