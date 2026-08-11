using System.Globalization;
using System.Text;

namespace OfficeIMO.Html;

internal static class HtmlCounterStyleFormatter {
    internal static bool TryFormat(int value, string style, out string formatted) {
        string decoded = HtmlCssEscapeDecoder.Decode(style.Trim());
        if (TryUnquote(decoded, out formatted)) return true;
        string normalized = decoded.ToLowerInvariant();
        switch (normalized) {
            case "decimal-leading-zero":
                formatted = value >= -9 && value <= 9
                    ? value < 0 ? "-0" + (-value).ToString(CultureInfo.InvariantCulture) : "0" + value.ToString(CultureInfo.InvariantCulture)
                    : value.ToString(CultureInfo.InvariantCulture);
                return true;
            case "lower-alpha":
            case "lower-latin":
                formatted = value > 0 ? FormatAlphabetic(value, "abcdefghijklmnopqrstuvwxyz") : value.ToString(CultureInfo.InvariantCulture);
                return true;
            case "upper-alpha":
            case "upper-latin":
                formatted = value > 0 ? FormatAlphabetic(value, "ABCDEFGHIJKLMNOPQRSTUVWXYZ") : value.ToString(CultureInfo.InvariantCulture);
                return true;
            case "lower-greek":
                formatted = value > 0 ? FormatAlphabetic(value, "αβγδεζηθικλμνξοπρστυφχψω") : value.ToString(CultureInfo.InvariantCulture);
                return true;
            case "lower-roman":
                formatted = value > 0 && value <= 3999 ? FormatRoman(value).ToLowerInvariant() : value.ToString(CultureInfo.InvariantCulture);
                return true;
            case "upper-roman":
                formatted = value > 0 && value <= 3999 ? FormatRoman(value) : value.ToString(CultureInfo.InvariantCulture);
                return true;
            case "disc":
                formatted = "•";
                return true;
            case "circle":
                formatted = "◦";
                return true;
            case "square":
                formatted = "▪";
                return true;
            case "none":
                formatted = string.Empty;
                return true;
            case "decimal":
            case "":
                formatted = value.ToString(CultureInfo.InvariantCulture);
                return true;
            default:
                formatted = string.Empty;
                return false;
        }
    }

    private static bool TryUnquote(string value, out string result) {
        if (value.Length >= 2 && (value[0] == '\'' && value[value.Length - 1] == '\'' || value[0] == '"' && value[value.Length - 1] == '"')) {
            result = value.Substring(1, value.Length - 2);
            return true;
        }
        result = string.Empty;
        return false;
    }

    private static string FormatAlphabetic(int value, string alphabet) {
        var result = new StringBuilder();
        int remaining = value;
        while (remaining > 0) {
            remaining--;
            result.Insert(0, alphabet[remaining % alphabet.Length]);
            remaining /= alphabet.Length;
        }
        return result.ToString();
    }

    private static string FormatRoman(int value) {
        var result = new StringBuilder();
        int[] values = { 1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1 };
        string[] symbols = { "M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I" };
        for (int index = 0; index < values.Length; index++) {
            while (value >= values[index]) {
                result.Append(symbols[index]);
                value -= values[index];
            }
        }
        return result.ToString();
    }
}
