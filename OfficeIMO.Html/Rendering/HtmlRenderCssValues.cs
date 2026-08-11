using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal static class HtmlRenderCssValues {
    internal static bool TryLength(string? value, double reference, double fontSize, double rootFontSize, out double result) {
        return TryLength(value, reference, fontSize, rootFontSize, double.NaN, double.NaN, out result);
    }

    internal static bool TryLength(
        string? value,
        double reference,
        double fontSize,
        double rootFontSize,
        double viewportWidth,
        double viewportHeight,
        out double result) {
        return TryLength(value, reference, fontSize, rootFontSize, viewportWidth, viewportHeight, double.NaN, double.NaN, out result);
    }

    internal static bool TryLength(
        string? value,
        double reference,
        double fontSize,
        double rootFontSize,
        double viewportWidth,
        double viewportHeight,
        double containerWidth,
        double containerHeight,
        out double result) {
        result = 0D;
        if (string.IsNullOrWhiteSpace(value)) {
            return false;
        }

        string normalized = value!.Trim().ToLowerInvariant();
        if (normalized == "0") {
            return true;
        }

        if (normalized == "auto" || normalized == "none") {
            return false;
        }

        if (normalized.IndexOf('(') >= 0) {
            return HtmlCssLengthMathEvaluator.TryEvaluate(normalized, reference, fontSize, rootFontSize, viewportWidth, viewportHeight, containerWidth, containerHeight, out result);
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
            case "vw":
            case "svw":
            case "lvw":
            case "dvw":
                result = number * viewportWidth / 100D;
                return IsFinite(result);
            case "vh":
            case "svh":
            case "lvh":
            case "dvh":
                result = number * viewportHeight / 100D;
                return IsFinite(result);
            case "vmin":
            case "svmin":
            case "lvmin":
            case "dvmin":
                result = number * Math.Min(viewportWidth, viewportHeight) / 100D;
                return IsFinite(result);
            case "vmax":
            case "svmax":
            case "lvmax":
            case "dvmax":
                result = number * Math.Max(viewportWidth, viewportHeight) / 100D;
                return IsFinite(result);
            case "cqw":
            case "cqi":
                result = number * (IsFinite(containerWidth) ? containerWidth : viewportWidth) / 100D;
                return IsFinite(result);
            case "cqh":
            case "cqb":
                result = number * (IsFinite(containerHeight) ? containerHeight : viewportHeight) / 100D;
                return IsFinite(result);
            case "cqmin":
                result = number * Math.Min(
                    IsFinite(containerWidth) ? containerWidth : viewportWidth,
                    IsFinite(containerHeight) ? containerHeight : viewportHeight) / 100D;
                return IsFinite(result);
            case "cqmax":
                result = number * Math.Max(
                    IsFinite(containerWidth) ? containerWidth : viewportWidth,
                    IsFinite(containerHeight) ? containerHeight : viewportHeight) / 100D;
                return IsFinite(result);
            case "%":
                result = reference * number / 100D;
                return IsFinite(result);
            default:
                return false;
        }
    }

    internal static bool HasExplicitLengthSyntax(string? value, bool allowPercentage, bool allowUnitlessZero) {
        if (string.IsNullOrWhiteSpace(value)) return false;

        string normalized = value!.Trim().ToLowerInvariant();
        if (normalized == "0") return allowUnitlessZero;
        if (!allowPercentage && normalized.IndexOf('%') >= 0) return false;
        if (normalized.IndexOf('(') >= 0) return true;

        int unitStart = normalized.Length;
        while (unitStart > 0 && (char.IsLetter(normalized[unitStart - 1]) || normalized[unitStart - 1] == '%')) {
            unitStart--;
        }
        return unitStart > 0 && unitStart < normalized.Length;
    }

    private static bool IsFinite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);

    internal static void ApplyBoxShorthand(string? value, double reference, double fontSize, double rootFontSize, ref double top, ref double right, ref double bottom, ref double left) {
        ApplyBoxShorthand(value, reference, fontSize, rootFontSize, double.NaN, double.NaN, ref top, ref right, ref bottom, ref left);
    }

    internal static void ApplyBoxShorthand(
        string? value,
        double reference,
        double fontSize,
        double rootFontSize,
        double viewportWidth,
        double viewportHeight,
        ref double top,
        ref double right,
        ref double bottom,
        ref double left) {
        IReadOnlyList<string> parts = SplitWhitespace(value);
        if (parts.Count == 0 || parts.Count > 4) {
            return;
        }

        var values = new double[parts.Count];
        for (int i = 0; i < parts.Count; i++) {
            if (!TryLength(parts[i], reference, fontSize, rootFontSize, viewportWidth, viewportHeight, out values[i])) {
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

        if (OfficeColor.TryParseCss(normalized, out color)) {
            return true;
        }

        IReadOnlyList<string> parts = SplitWhitespace(normalized);
        for (int i = parts.Count - 1; i >= 0; i--) {
            if (OfficeColor.TryParseCss(parts[i].Trim(',', ';'), out color)) {
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

    internal static int FindMatchingParenthesis(string text, int openIndex) {
        if (openIndex < 0 || openIndex >= text.Length || text[openIndex] != '(') return -1;

        int depth = 0;
        char quote = '\0';
        for (int index = openIndex; index < text.Length; index++) {
            char current = text[index];
            if (current == '\\') {
                if (HtmlCssEscapeDecoder.TryDecodeEscape(text, index, out _, out int consumedCharacters)) {
                    index += consumedCharacters - 1;
                } else if (index + 1 < text.Length) {
                    index++;
                }
                continue;
            }
            if (quote != '\0') {
                if (current == quote) quote = '\0';
                continue;
            }
            if (current == '\'' || current == '"') {
                quote = current;
            } else if (current == '(') {
                depth++;
            } else if (current == ')' && --depth == 0) {
                return index;
            }
        }

        return -1;
    }

    internal static bool TrySplitTopLevelCommas(string? value, int maximumParts, out IReadOnlyList<string> parts) {
        parts = Array.Empty<string>();
        if (maximumParts <= 0 || string.IsNullOrWhiteSpace(value)) return false;

        var resolved = new List<string>(Math.Min(maximumParts, 16));
        int start = 0;
        int depth = 0;
        char quote = '\0';
        string text = value!;
        for (int index = 0; index < text.Length; index++) {
            char current = text[index];
            if (current == '\\') {
                if (HtmlCssEscapeDecoder.TryDecodeEscape(text, index, out _, out int consumedCharacters)) {
                    index += consumedCharacters - 1;
                } else if (index + 1 < text.Length) {
                    index++;
                }
                continue;
            }
            if (quote != '\0') {
                if (current == quote) quote = '\0';
                continue;
            }

            if (current == '\'' || current == '"') quote = current;
            else if (current == '(') depth++;
            else if (current == ')' && depth > 0) depth--;
            else if (current == ',' && depth == 0) {
                if (resolved.Count >= maximumParts) return false;
                resolved.Add(text.Substring(start, index - start).Trim());
                start = index + 1;
            }
        }

        if (resolved.Count >= maximumParts) return false;
        resolved.Add(text.Substring(start).Trim());
        parts = resolved.AsReadOnly();
        return true;
    }

    internal static IReadOnlyList<string> SplitTopLevel(string? value, char separator) {
        if (string.IsNullOrWhiteSpace(value)) return Array.Empty<string>();

        var parts = new List<string>();
        int start = 0;
        int depth = 0;
        char quote = '\0';
        string text = value!;
        for (int index = 0; index < text.Length; index++) {
            char current = text[index];
            if (current == '\\') {
                if (HtmlCssEscapeDecoder.TryDecodeEscape(text, index, out _, out int consumedCharacters)) {
                    index += consumedCharacters - 1;
                } else if (index + 1 < text.Length) {
                    index++;
                }
                continue;
            }
            if (quote != '\0') {
                if (current == quote) quote = '\0';
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

}
