using System;

namespace OfficeIMO.Drawing;

/// <summary>
/// Applies the color transformations shared by Office document formats.
/// </summary>
public static class OfficeColorTransforms {
    /// <summary>
    /// Mixes the input color with white. <paramref name="inputRatio"/> is the proportion of the
    /// original color retained, matching the DrawingML <c>tint</c> contract.
    /// </summary>
    public static OfficeColor Tint(OfficeColor color, double inputRatio) {
        ValidateUnitInterval(inputRatio, nameof(inputRatio));
        return OfficeColor.FromRgba(
            Mix(color.R, 255, inputRatio),
            Mix(color.G, 255, inputRatio),
            Mix(color.B, 255, inputRatio),
            color.A);
    }

    /// <summary>
    /// Mixes the input color with black. <paramref name="inputRatio"/> is the proportion of the
    /// original color retained, matching the DrawingML <c>shade</c> contract.
    /// </summary>
    public static OfficeColor Shade(OfficeColor color, double inputRatio) {
        ValidateUnitInterval(inputRatio, nameof(inputRatio));
        return OfficeColor.FromRgba(
            ScaleChannel(color.R, inputRatio),
            ScaleChannel(color.G, inputRatio),
            ScaleChannel(color.B, inputRatio),
            color.A);
    }

    /// <summary>Replaces the color alpha with a normalized value from 0 to 1.</summary>
    public static OfficeColor WithAlpha(OfficeColor color, double alpha) {
        ValidateUnitInterval(alpha, nameof(alpha));
        return OfficeColor.FromRgba(color.R, color.G, color.B, ToChannel(alpha * 255D));
    }

    /// <summary>Multiplies the color alpha and clamps the result to the valid channel range.</summary>
    public static OfficeColor ModulateAlpha(OfficeColor color, double factor) {
        ValidateNonNegativeFinite(factor, nameof(factor));
        return OfficeColor.FromRgba(color.R, color.G, color.B, ToChannel(color.A * factor));
    }

    /// <summary>Adds a normalized offset to the color alpha and clamps the result.</summary>
    public static OfficeColor OffsetAlpha(OfficeColor color, double offset) {
        ValidateFinite(offset, nameof(offset));
        return OfficeColor.FromRgba(color.R, color.G, color.B, ToChannel(color.A + (255D * offset)));
    }

    /// <summary>Multiplies HSL luminance and clamps the result to the normalized range.</summary>
    public static OfficeColor ModulateLuminance(OfficeColor color, double factor) {
        ValidateNonNegativeFinite(factor, nameof(factor));
        ToHsl(color, out double hue, out double saturation, out double luminance);
        return FromHsl(hue, saturation, ClampUnit(luminance * factor), color.A);
    }

    /// <summary>Adds a normalized HSL luminance offset and clamps the result.</summary>
    public static OfficeColor OffsetLuminance(OfficeColor color, double offset) {
        ValidateFinite(offset, nameof(offset));
        ToHsl(color, out double hue, out double saturation, out double luminance);
        return FromHsl(hue, saturation, ClampUnit(luminance + offset), color.A);
    }

    /// <summary>
    /// Applies a SpreadsheetML tint in the inclusive range -1 to 1 by adjusting HSL luminance.
    /// Negative values darken and positive values lighten.
    /// </summary>
    public static OfficeColor SpreadsheetTint(OfficeColor color, double tint) {
        ValidateFinite(tint, nameof(tint));
        if (tint < -1D || tint > 1D) {
            throw new ArgumentOutOfRangeException(nameof(tint), "Spreadsheet tint must be between -1 and 1.");
        }

        ToHsl(color, out double hue, out double saturation, out double luminance);
        double transformed = tint < 0D
            ? luminance * (1D + tint)
            : luminance * (1D - tint) + tint;
        return FromHsl(hue, saturation, ClampUnit(transformed), color.A);
    }

    private static void ToHsl(OfficeColor color, out double hue, out double saturation, out double luminance) {
        double red = color.R / 255D;
        double green = color.G / 255D;
        double blue = color.B / 255D;
        double maximum = Math.Max(red, Math.Max(green, blue));
        double minimum = Math.Min(red, Math.Min(green, blue));
        double delta = maximum - minimum;

        luminance = (maximum + minimum) / 2D;
        if (delta <= 0.0000000001D) {
            hue = 0D;
            saturation = 0D;
            return;
        }

        saturation = delta / (1D - Math.Abs((2D * luminance) - 1D));
        if (maximum == red) {
            hue = ((green - blue) / delta) % 6D;
        } else if (maximum == green) {
            hue = ((blue - red) / delta) + 2D;
        } else {
            hue = ((red - green) / delta) + 4D;
        }

        hue *= 60D;
        if (hue < 0D) {
            hue += 360D;
        }
    }

    private static OfficeColor FromHsl(double hue, double saturation, double luminance, byte alpha) {
        hue = ((hue % 360D) + 360D) % 360D;
        saturation = ClampUnit(saturation);
        luminance = ClampUnit(luminance);

        double chroma = (1D - Math.Abs((2D * luminance) - 1D)) * saturation;
        double sector = hue / 60D;
        double secondary = chroma * (1D - Math.Abs((sector % 2D) - 1D));
        double red;
        double green;
        double blue;
        if (sector < 1D) {
            red = chroma; green = secondary; blue = 0D;
        } else if (sector < 2D) {
            red = secondary; green = chroma; blue = 0D;
        } else if (sector < 3D) {
            red = 0D; green = chroma; blue = secondary;
        } else if (sector < 4D) {
            red = 0D; green = secondary; blue = chroma;
        } else if (sector < 5D) {
            red = secondary; green = 0D; blue = chroma;
        } else {
            red = chroma; green = 0D; blue = secondary;
        }

        double match = luminance - (chroma / 2D);
        return OfficeColor.FromRgba(
            ToChannel((red + match) * 255D),
            ToChannel((green + match) * 255D),
            ToChannel((blue + match) * 255D),
            alpha);
    }

    private static byte Mix(byte input, byte other, double inputRatio) =>
        ToChannel((input * inputRatio) + (other * (1D - inputRatio)));

    private static byte ScaleChannel(byte channel, double factor) => ToChannel(channel * factor);

    private static byte ToChannel(double value) =>
        (byte)Math.Max(0D, Math.Min(255D, Math.Round(value, MidpointRounding.ToEven)));

    private static double ClampUnit(double value) => Math.Max(0D, Math.Min(1D, value));

    private static void ValidateUnitInterval(double value, string parameterName) {
        ValidateFinite(value, parameterName);
        if (value < 0D || value > 1D) {
            throw new ArgumentOutOfRangeException(parameterName, "Value must be between 0 and 1.");
        }
    }

    private static void ValidateNonNegativeFinite(double value, string parameterName) {
        ValidateFinite(value, parameterName);
        if (value < 0D) {
            throw new ArgumentOutOfRangeException(parameterName, "Value cannot be negative.");
        }
    }

    private static void ValidateFinite(double value, string parameterName) {
        if (double.IsNaN(value) || double.IsInfinity(value)) {
            throw new ArgumentOutOfRangeException(parameterName, "Value must be finite.");
        }
    }
}
