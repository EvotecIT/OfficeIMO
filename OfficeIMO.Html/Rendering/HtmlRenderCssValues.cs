using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal static class HtmlRenderCssValues {
    internal static bool TryLength(string? value, double reference, double fontSize, double rootFontSize, out double result) {
        result = 0D;
        if (string.IsNullOrWhiteSpace(value)) {
            return false;
        }

        string normalized = value!.Trim().ToLowerInvariant();
        if (normalized == "0") {
            return true;
        }

        if (normalized == "auto" || normalized == "none" || normalized.IndexOf('(') >= 0) {
            return false;
        }

        string unit = string.Empty;
        int unitStart = normalized.Length;
        while (unitStart > 0 && (char.IsLetter(normalized[unitStart - 1]) || normalized[unitStart - 1] == '%')) {
            unitStart--;
        }

        if (unitStart < normalized.Length) {
            unit = normalized.Substring(unitStart);
            normalized = normalized.Substring(0, unitStart).Trim();
        }

        if (!double.TryParse(normalized, NumberStyles.Float, CultureInfo.InvariantCulture, out double number)
            || double.IsNaN(number)
            || double.IsInfinity(number)) {
            return false;
        }

        switch (unit) {
            case "":
            case "px":
                result = number;
                return IsFinite(result);
            case "pt":
                result = number * HtmlRenderOptions.CssPixelsPerInch / 72D;
                return IsFinite(result);
            case "pc":
                result = number * HtmlRenderOptions.CssPixelsPerInch / 6D;
                return IsFinite(result);
            case "in":
                result = number * HtmlRenderOptions.CssPixelsPerInch;
                return IsFinite(result);
            case "cm":
                result = number * HtmlRenderOptions.CssPixelsPerInch / 2.54D;
                return IsFinite(result);
            case "mm":
                result = number * HtmlRenderOptions.CssPixelsPerInch / 25.4D;
                return IsFinite(result);
            case "q":
                result = number * HtmlRenderOptions.CssPixelsPerInch / 101.6D;
                return IsFinite(result);
            case "em":
                result = number * fontSize;
                return IsFinite(result);
            case "rem":
                result = number * rootFontSize;
                return IsFinite(result);
            case "%":
                result = reference * number / 100D;
                return IsFinite(result);
            default:
                return false;
        }
    }

    private static bool IsFinite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);

    internal static void ApplyBoxShorthand(string? value, double reference, double fontSize, double rootFontSize, ref double top, ref double right, ref double bottom, ref double left) {
        IReadOnlyList<string> parts = SplitWhitespace(value);
        if (parts.Count == 0 || parts.Count > 4) {
            return;
        }

        var values = new double[parts.Count];
        for (int i = 0; i < parts.Count; i++) {
            if (!TryLength(parts[i], reference, fontSize, rootFontSize, out values[i])) {
                return;
            }
        }

        top = values[0];
        right = parts.Count > 1 ? values[1] : values[0];
        bottom = parts.Count > 2 ? values[2] : values[0];
        left = parts.Count > 3 ? values[3] : right;
    }

    internal static bool TryColor(string? value, out OfficeColor color) {
        color = default;
        if (string.IsNullOrWhiteSpace(value)) {
            return false;
        }

        string normalized = value!.Trim();
        if (string.Equals(normalized, "transparent", StringComparison.OrdinalIgnoreCase)) {
            color = OfficeColor.Transparent;
            return true;
        }

        if (OfficeColor.TryParse(normalized, out color)) {
            return true;
        }

        int rgbStart = normalized.IndexOf("rgb", StringComparison.OrdinalIgnoreCase);
        if (rgbStart >= 0) {
            int open = normalized.IndexOf('(', rgbStart);
            int close = open >= 0 ? normalized.IndexOf(')', open + 1) : -1;
            if (open >= 0 && close > open && TryRgbFunction(normalized.Substring(open + 1, close - open - 1), out color)) {
                return true;
            }
        }

        IReadOnlyList<string> parts = SplitWhitespace(normalized);
        for (int i = parts.Count - 1; i >= 0; i--) {
            if (OfficeColor.TryParse(parts[i].Trim(',', ';'), out color)) {
                return true;
            }
        }

        return false;
    }

    internal static string FontFamilyList(string? value, string fallback) {
        if (string.IsNullOrWhiteSpace(value)) {
            return fallback;
        }

        string normalized = value!.Trim();
        return normalized.Length == 0 ? fallback : normalized;
    }

    internal static IReadOnlyList<string> SplitWhitespace(string? value) {
        if (string.IsNullOrWhiteSpace(value)) {
            return Array.Empty<string>();
        }

        var parts = new List<string>();
        int start = -1;
        int depth = 0;
        int bracketDepth = 0;
        char quote = '\0';
        string text = value!;
        for (int i = 0; i < text.Length; i++) {
            char current = text[i];
            if (quote != '\0') {
                if (current == quote && (i == 0 || text[i - 1] != '\\')) {
                    quote = '\0';
                }

                continue;
            }

            if (current == '\'' || current == '"') {
                quote = current;
                if (start < 0) start = i;
                continue;
            }

            if (current == '(') depth++;
            if (current == ')' && depth > 0) depth--;
            if (current == '[') bracketDepth++;
            if (current == ']' && bracketDepth > 0) bracketDepth--;
            if (char.IsWhiteSpace(current) && depth == 0 && bracketDepth == 0) {
                if (start >= 0) {
                    parts.Add(text.Substring(start, i - start));
                    start = -1;
                }

                continue;
            }

            if (start < 0) start = i;
        }

        if (start >= 0) {
            parts.Add(text.Substring(start));
        }

        return parts;
    }

    internal static IReadOnlyList<string> SplitTopLevelCommas(string? value) => SplitTopLevel(value, ',');

    internal static IReadOnlyList<string> SplitTopLevel(string? value, char separator) {
        if (string.IsNullOrWhiteSpace(value)) return Array.Empty<string>();

        var parts = new List<string>();
        int start = 0;
        int depth = 0;
        char quote = '\0';
        string text = value!;
        for (int index = 0; index < text.Length; index++) {
            char current = text[index];
            if (quote != '\0') {
                if (current == quote && (index == 0 || text[index - 1] != '\\')) quote = '\0';
                continue;
            }

            if (current == '\'' || current == '"') {
                quote = current;
            } else if (current == '(') {
                depth++;
            } else if (current == ')' && depth > 0) {
                depth--;
            } else if (current == separator && depth == 0) {
                parts.Add(text.Substring(start, index - start).Trim());
                start = index + 1;
            }
        }

        parts.Add(text.Substring(start).Trim());
        return parts.AsReadOnly();
    }

    internal static OfficeColor ApplyOpacity(OfficeColor color, double opacity) {
        if (opacity >= 1D) return color;
        if (opacity <= 0D) return OfficeColor.FromRgba(color.R, color.G, color.B, 0);
        return OfficeColor.FromRgba(color.R, color.G, color.B, (byte)Math.Round(color.A * opacity));
    }

    private static bool TryRgbFunction(string arguments, out OfficeColor color) {
        color = default;
        string normalized = arguments.Replace('/', ',');
        string[] parts = normalized.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length == 1) {
            parts = normalized.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
        }

        if (parts.Length < 3 || !TryColorChannel(parts[0], out byte red) || !TryColorChannel(parts[1], out byte green) || !TryColorChannel(parts[2], out byte blue)) {
            return false;
        }

        byte alpha = 255;
        if (parts.Length > 3 && !TryAlpha(parts[3], out alpha)) {
            return false;
        }

        color = OfficeColor.FromRgba(red, green, blue, alpha);
        return true;
    }

    private static bool TryColorChannel(string value, out byte channel) {
        channel = 0;
        string normalized = value.Trim();
        bool percent = normalized.EndsWith("%", StringComparison.Ordinal);
        if (percent) normalized = normalized.Substring(0, normalized.Length - 1);
        if (!double.TryParse(normalized, NumberStyles.Float, CultureInfo.InvariantCulture, out double number)) {
            return false;
        }

        if (percent) number = number * 255D / 100D;
        number = Math.Max(0D, Math.Min(255D, number));
        channel = (byte)Math.Round(number);
        return true;
    }

    private static bool TryAlpha(string value, out byte alpha) {
        alpha = 255;
        string normalized = value.Trim();
        bool percent = normalized.EndsWith("%", StringComparison.Ordinal);
        if (percent) normalized = normalized.Substring(0, normalized.Length - 1);
        if (!double.TryParse(normalized, NumberStyles.Float, CultureInfo.InvariantCulture, out double number)) {
            return false;
        }

        if (percent) number /= 100D;
        number = Math.Max(0D, Math.Min(1D, number));
        alpha = (byte)Math.Round(number * 255D);
        return true;
    }
}
