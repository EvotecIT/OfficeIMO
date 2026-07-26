using System;
using System.Globalization;

namespace OfficeIMO.Drawing;

public readonly partial struct OfficeColor {
    private const int MaximumCssColorLength = 256;

    /// <summary>
    /// Parses a CSS named, hexadecimal, rgb(), rgba(), hsl(), or hsla() color.
    /// Both legacy comma syntax and modern space/slash syntax are accepted.
    /// </summary>
    public static OfficeColor ParseCss(string value) {
        if (TryParseCss(value, out OfficeColor color)) return color;
        throw new FormatException($"Invalid CSS color value: '{value}'.");
    }

    /// <summary>
    /// Tries to parse a CSS named, hexadecimal, rgb(), rgba(), hsl(), or hsla() color.
    /// Out-of-range numeric channels are clamped as required by CSS color serialization.
    /// </summary>
    public static bool TryParseCss(string? value, out OfficeColor color) {
        color = default;
        if (string.IsNullOrWhiteSpace(value) || value!.Length > MaximumCssColorLength) return false;

        string normalized = value.Trim();
        if (normalized[0] == '#') return TryParseHex(normalized, out color);
        if (NamedColors.TryGetValue(normalized, out color)) return true;
        if (!TryReadFunction(normalized, out string name, out string arguments)) return false;

        if (name == "rgb" || name == "rgba") {
            return TryParseRgb(arguments, out color);
        }
        if (name == "hsl" || name == "hsla") {
            return TryParseHsl(arguments, out color);
        }
        return false;
    }

    private static bool TryReadFunction(string value, out string name, out string arguments) {
        name = string.Empty;
        arguments = string.Empty;
        int open = value.IndexOf('(');
        if (open <= 0 || value[value.Length - 1] != ')' || value.IndexOf('(', open + 1) >= 0) return false;
        name = value.Substring(0, open).Trim().ToLowerInvariant();
        arguments = value.Substring(open + 1, value.Length - open - 2).Trim();
        return arguments.Length > 0 && arguments.IndexOf(')') < 0;
    }

    private static bool TryParseRgb(string arguments, out OfficeColor color) {
        color = default;
        if (!TrySplitFunctionArguments(arguments, out string[] channels, out string? alpha)
            || channels.Length != 3
            || !TryRgbChannel(channels[0], out byte red)
            || !TryRgbChannel(channels[1], out byte green)
            || !TryRgbChannel(channels[2], out byte blue)
            || !TryAlphaChannel(alpha, out byte opacity)) {
            return false;
        }

        color = FromRgba(red, green, blue, opacity);
        return true;
    }

    private static bool TryParseHsl(string arguments, out OfficeColor color) {
        color = default;
        if (!TrySplitFunctionArguments(arguments, out string[] channels, out string? alpha)
            || channels.Length != 3
            || !TryHue(channels[0], out double hue)
            || !TryPercentage(channels[1], out double saturation)
            || !TryPercentage(channels[2], out double lightness)
            || !TryAlphaChannel(alpha, out byte opacity)) {
            return false;
        }

        double chroma = (1D - Math.Abs((2D * lightness) - 1D)) * saturation;
        double sector = hue / 60D;
        double secondary = chroma * (1D - Math.Abs((sector % 2D) - 1D));
        double red = 0D;
        double green = 0D;
        double blue = 0D;
        if (sector < 1D) {
            red = chroma;
            green = secondary;
        } else if (sector < 2D) {
            red = secondary;
            green = chroma;
        } else if (sector < 3D) {
            green = chroma;
            blue = secondary;
        } else if (sector < 4D) {
            green = secondary;
            blue = chroma;
        } else if (sector < 5D) {
            red = secondary;
            blue = chroma;
        } else {
            red = chroma;
            blue = secondary;
        }

        double match = lightness - (chroma / 2D);
        color = FromRgba(
            ToByte((red + match) * 255D),
            ToByte((green + match) * 255D),
            ToByte((blue + match) * 255D),
            opacity);
        return true;
    }

    private static bool TrySplitFunctionArguments(string arguments, out string[] channels, out string? alpha) {
        channels = Array.Empty<string>();
        alpha = null;
        if (arguments.IndexOf(',') >= 0) {
            string[] parts = arguments.Split(',');
            if (parts.Length != 3 && parts.Length != 4) return false;
            channels = new[] { parts[0].Trim(), parts[1].Trim(), parts[2].Trim() };
            alpha = parts.Length == 4 ? parts[3].Trim() : null;
            return channels[0].Length > 0 && channels[1].Length > 0 && channels[2].Length > 0
                && (alpha == null || alpha.Length > 0);
        }

        string[] slashParts = arguments.Split('/');
        if (slashParts.Length > 2) return false;
        channels = slashParts[0].Split(
            new[] { ' ', '\t', '\r', '\n', '\f' },
            StringSplitOptions.RemoveEmptyEntries);
        alpha = slashParts.Length == 2 ? slashParts[1].Trim() : null;
        return channels.Length == 3 && (alpha == null || alpha.Length > 0);
    }

    private static bool TryRgbChannel(string value, out byte channel) {
        channel = 0;
        bool percentage = value.EndsWith("%", StringComparison.Ordinal);
        string numberText = percentage ? value.Substring(0, value.Length - 1).Trim() : value.Trim();
        if (!TryFiniteDouble(numberText, out double number)) return false;
        channel = ToByte(percentage ? number * 255D / 100D : number);
        return true;
    }

    private static bool TryAlphaChannel(string? value, out byte alpha) {
        alpha = 255;
        if (value == null) return true;
        bool percentage = value.EndsWith("%", StringComparison.Ordinal);
        string numberText = percentage ? value.Substring(0, value.Length - 1).Trim() : value.Trim();
        if (!TryFiniteDouble(numberText, out double number)) return false;
        alpha = ToByte((percentage ? number / 100D : number) * 255D);
        return true;
    }

    private static bool TryHue(string value, out double degrees) {
        degrees = 0D;
        string normalized = value.Trim().ToLowerInvariant();
        double multiplier = 1D;
        if (normalized.EndsWith("grad", StringComparison.Ordinal)) {
            normalized = normalized.Substring(0, normalized.Length - 4).Trim();
            multiplier = 0.9D;
        } else if (normalized.EndsWith("turn", StringComparison.Ordinal)) {
            normalized = normalized.Substring(0, normalized.Length - 4).Trim();
            multiplier = 360D;
        } else if (normalized.EndsWith("rad", StringComparison.Ordinal)) {
            normalized = normalized.Substring(0, normalized.Length - 3).Trim();
            multiplier = 180D / Math.PI;
        } else if (normalized.EndsWith("deg", StringComparison.Ordinal)) {
            normalized = normalized.Substring(0, normalized.Length - 3).Trim();
        }
        if (!TryFiniteDouble(normalized, out double number)) return false;
        degrees = (number * multiplier) % 360D;
        if (degrees < 0D) degrees += 360D;
        return true;
    }

    private static bool TryPercentage(string value, out double fraction) {
        fraction = 0D;
        string normalized = value.Trim();
        if (!normalized.EndsWith("%", StringComparison.Ordinal)
            || !TryFiniteDouble(normalized.Substring(0, normalized.Length - 1).Trim(), out double number)) {
            return false;
        }
        fraction = Clamp(number, 0D, 100D) / 100D;
        return true;
    }

    private static bool TryFiniteDouble(string value, out double number) =>
        double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out number)
        && !double.IsNaN(number)
        && !double.IsInfinity(number);

    private static byte ToByte(double value) =>
        (byte)Math.Round(Clamp(value, 0D, 255D));

    private static double Clamp(double value, double minimum, double maximum) =>
        value < minimum ? minimum : value > maximum ? maximum : value;
}
