using OfficeIMO.Drawing;

namespace OfficeIMO.OpenXml.Internal;

/// <summary>
/// The DrawingML color arithmetic: parsing the color models and applying the ordered color transforms. It works on
/// element names and attribute values, so every XML layer that reads DrawingML colors shares it.
/// </summary>
internal struct OfficeDrawingColorTransforms {
    private OfficeColor _resolved;
    private bool _pendingHsl;
    private double _hue, _saturation, _luminance;

    internal OfficeDrawingColorTransforms(OfficeColor color) {
        _resolved = color;
        _pendingHsl = false;
        _hue = _saturation = _luminance = 0D;
    }

    /// <summary>Whether the transform reads a <c>val</c> attribute; <c>comp</c>, <c>inv</c> and <c>gray</c> do not.</summary>
    internal static bool TakesValue(string name) => name is not ("comp" or "inv" or "gray");

    /// <summary>The names of the transforms <see cref="Apply"/> knows.</summary>
    internal static bool IsKnown(string name) => name is "comp" or "inv" or "gray" or "alpha" or "alphaMod" or "alphaOff" or "tint" or "shade" or
        "lumMod" or "lumOff" or "hueOff" or "hueMod" or "satOff" or "satMod" or "red" or "redMod" or "redOff" or "green" or "greenMod" or
        "greenOff" or "blue" or "blueMod" or "blueOff";

    /// <summary>The color after the transforms applied so far.</summary>
    internal OfficeColor Result => _pendingHsl ? OfficeColorTransforms.FromHsl(_hue, _saturation, _luminance, _resolved.A) : _resolved;

    /// <summary>
    /// Applies one transform, in child order, retaining each transform's alpha semantics. A transform with a missing or
    /// unreadable value, or an unknown name, changes nothing. Tint and shade retain the specified input proportion in
    /// linear sRGB.
    /// </summary>
    internal void Apply(string name, string? rawValue) {
        if (name is "comp" or "lumMod" or "lumOff" or "hueOff" or "hueMod" or "satOff" or "satMod") {
            double amount = 0D;
            if (name != "comp" && !TryReadValue(rawValue, out amount)) return;
            if (!_pendingHsl) {
                OfficeColorTransforms.ToHsl(_resolved, out _hue, out _saturation, out _luminance);
                _pendingHsl = true;
            }
            switch (name) {
                case "comp": _hue = (_hue + 180D) % 360D; break;
                case "lumMod": _luminance = ClampUnit(_luminance * Math.Max(0D, amount)); break;
                case "lumOff": _luminance = ClampUnit(_luminance + amount); break;
                case "hueOff": _hue = ((_hue + amount * 100000D / 60000D) % 360D + 360D) % 360D; break;
                case "hueMod": _hue = _hue * Math.Max(0D, amount) % 360D; break;
                case "satOff": _saturation = ClampUnit(_saturation + amount); break;
                case "satMod": _saturation = ClampUnit(_saturation * Math.Max(0D, amount)); break;
            }
            return;
        }
        // Alpha does not change the HSL channels. Other color-space operations must observe the completed HSL
        // sequence; do not round between HSL modifiers.
        if (_pendingHsl && name is not ("alpha" or "alphaMod" or "alphaOff")) {
            _resolved = OfficeColorTransforms.FromHsl(_hue, _saturation, _luminance, _resolved.A);
            _pendingHsl = false;
        }
        OfficeColor resolved = _resolved;
        switch (name) {
            case "inv":
                _resolved = OfficeColor.FromRgba((byte)(255 - resolved.R), (byte)(255 - resolved.G), (byte)(255 - resolved.B), resolved.A);
                return;
            case "gray":
                byte gray = ToChannel((resolved.R * 0.299D) + (resolved.G * 0.587D) + (resolved.B * 0.114D));
                _resolved = OfficeColor.FromRgba(gray, gray, gray, resolved.A);
                return;
        }
        if (!TryReadValue(rawValue, out double value)) return;
        _resolved = name switch {
            "alpha" => OfficeColorTransforms.WithAlpha(resolved, ClampUnit(value)),
            "alphaMod" => OfficeColorTransforms.ModulateAlpha(resolved, Math.Max(0D, value)),
            "alphaOff" => OfficeColorTransforms.OffsetAlpha(resolved, value),
            "tint" => OfficeGradientColors.Interpolate(OfficeColor.FromRgba(255, 255, 255, resolved.A), resolved, ClampUnit(value), OfficeGradientColorInterpolation.LinearRgb),
            "shade" => OfficeGradientColors.Interpolate(OfficeColor.FromRgba(0, 0, 0, resolved.A), resolved, ClampUnit(value), OfficeGradientColorInterpolation.LinearRgb),
            "red" => OfficeColor.FromRgba(ToChannel(255D * value), resolved.G, resolved.B, resolved.A),
            "redMod" => OfficeColor.FromRgba(ToChannel(resolved.R * value), resolved.G, resolved.B, resolved.A),
            "redOff" => OfficeColor.FromRgba(ToChannel(resolved.R + (255D * value)), resolved.G, resolved.B, resolved.A),
            "green" => OfficeColor.FromRgba(resolved.R, ToChannel(255D * value), resolved.B, resolved.A),
            "greenMod" => OfficeColor.FromRgba(resolved.R, ToChannel(resolved.G * value), resolved.B, resolved.A),
            "greenOff" => OfficeColor.FromRgba(resolved.R, ToChannel(resolved.G + (255D * value)), resolved.B, resolved.A),
            "blue" => OfficeColor.FromRgba(resolved.R, resolved.G, ToChannel(255D * value), resolved.A),
            "blueMod" => OfficeColor.FromRgba(resolved.R, resolved.G, ToChannel(resolved.B * value), resolved.A),
            "blueOff" => OfficeColor.FromRgba(resolved.R, resolved.G, ToChannel(resolved.B + (255D * value)), resolved.A),
            _ => resolved
        };
    }

    /// <summary>A transform value, stored in thousandths of a percent, as a fraction.</summary>
    internal static bool TryReadValue(string? raw, out double value) {
        value = 0D;
        if (ParsePercentage(raw) is not int scaled) return false;
        value = scaled / 100000D;
        return true;
    }

    /// <summary>
    /// A DrawingML percentage in thousandths of a percent: an integer as written, or a decimal with a trailing
    /// <c>%</c> (the form strict documents use), so <c>50000</c> and <c>50%</c> read the same.
    /// </summary>
    internal static int? ParsePercentage(string? raw) {
        if (string.IsNullOrWhiteSpace(raw)) return null;
        string text = raw!.Trim();
        if (int.TryParse(text, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int scaled)) return scaled;
        if (text.Length > 1 && text[text.Length - 1] == '%'
            && double.TryParse(text.Substring(0, text.Length - 1), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double percent)
            && !double.IsNaN(percent) && !double.IsInfinity(percent)) {
            double thousandths = Math.Round(percent * 1000D, MidpointRounding.AwayFromZero);
            if (thousandths >= int.MinValue && thousandths <= int.MaxValue) return (int)thousandths;
        }
        return null;
    }

    /// <summary>Reads the DrawingML token, whose dk/lt/med prefixes abbreviate CSS color names.</summary>
    internal static OfficeColor? ParsePreset(string? token) {
        if (string.IsNullOrWhiteSpace(token)) return null;
        string name = token!;
        if (name.StartsWith("dk", StringComparison.OrdinalIgnoreCase)) name = "dark" + name.Substring(2);
        else if (name.StartsWith("lt", StringComparison.OrdinalIgnoreCase)) name = "light" + name.Substring(2);
        else if (name.StartsWith("med", StringComparison.OrdinalIgnoreCase)) name = "medium" + name.Substring(3);
        return OfficeColor.TryParse(name, out OfficeColor color) ? color : (OfficeColor?)null;
    }

    internal static OfficeColor? ParseRgb(string? value) =>
        OfficeColor.TryParseHex(value, out OfficeColor color) ? color : (OfficeColor?)null;

    /// <summary>An <c>a:scrgbClr</c>, whose portions are linear light in thousandths of a percent.</summary>
    internal static OfficeColor? ParseScRgb(int? red, int? green, int? blue) {
        if (red is not int r || green is not int g || blue is not int b) return null;
        return OfficeColor.FromRgb(
            ToChannel(255D * LinearScRgbToSrgb(r / 100000D)),
            ToChannel(255D * LinearScRgbToSrgb(g / 100000D)),
            ToChannel(255D * LinearScRgbToSrgb(b / 100000D)));
    }

    private static double LinearScRgbToSrgb(double value) {
        double linear = ClampUnit(value);
        return linear <= 0.0031308D ? 12.92D * linear : 1.055D * Math.Pow(linear, 1D / 2.4D) - 0.055D;
    }

    /// <summary>An <c>a:hslClr</c>: hue in 60000ths of a degree, saturation and luminance in thousandths of a percent.</summary>
    internal static OfficeColor? ParseHsl(int? hueValue, int? saturationValue, int? luminanceValue) {
        if (hueValue is not int hue || saturationValue is not int saturation || luminanceValue is not int luminance) return null;
        double normalizedHue = ((hue / 60000D) % 360D + 360D) % 360D;
        double normalizedSaturation = ClampUnit(saturation / 100000D);
        double normalizedLuminance = ClampUnit(luminance / 100000D);
        double chroma = (1D - Math.Abs(2D * normalizedLuminance - 1D)) * normalizedSaturation;
        double sector = normalizedHue / 60D;
        double intermediate = chroma * (1D - Math.Abs(sector % 2D - 1D));
        double red, green, blue;
        if (sector < 1D) {
            red = chroma; green = intermediate; blue = 0D;
        } else if (sector < 2D) {
            red = intermediate; green = chroma; blue = 0D;
        } else if (sector < 3D) {
            red = 0D; green = chroma; blue = intermediate;
        } else if (sector < 4D) {
            red = 0D; green = intermediate; blue = chroma;
        } else if (sector < 5D) {
            red = intermediate; green = 0D; blue = chroma;
        } else {
            red = chroma; green = 0D; blue = intermediate;
        }
        double match = normalizedLuminance - chroma / 2D;
        return OfficeColor.FromRgb(ToChannel(255D * (red + match)), ToChannel(255D * (green + match)), ToChannel(255D * (blue + match)));
    }

    internal static byte ToChannel(double value) =>
        (byte)Math.Max(0D, Math.Min(255D, Math.Round(value, MidpointRounding.ToEven)));

    internal static double ClampUnit(double value) => Math.Max(0D, Math.Min(1D, value));
}
