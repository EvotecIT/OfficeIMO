using System.Globalization;

namespace OfficeIMO.Rtf.Html;

internal static class HtmlStyleDeclarationParser {
    internal static HtmlStyleDeclaration Parse(string? style) {
        if (string.IsNullOrWhiteSpace(style)) {
            return HtmlStyleDeclaration.Empty;
        }

        var declaration = new HtmlStyleDeclaration();
        foreach (string rawDeclaration in SplitDeclarations(style!)) {
            int separator = rawDeclaration.IndexOf(':');
            if (separator <= 0) {
                continue;
            }

            string property = rawDeclaration.Substring(0, separator).Trim().ToLowerInvariant();
            string value = NormalizeValue(rawDeclaration.Substring(separator + 1));
            if (value.Length == 0) {
                continue;
            }

            Apply(declaration, property, value);
        }

        return declaration;
    }

    private static void Apply(HtmlStyleDeclaration declaration, string property, string value) {
        switch (property) {
            case "font-weight":
                declaration.Bold = ParseFontWeight(value);
                break;
            case "font-style":
                declaration.Italic = ParseFontStyle(value);
                break;
            case "text-decoration":
            case "text-decoration-line":
                ApplyTextDecoration(declaration, value);
                break;
            case "vertical-align":
                declaration.VerticalPosition = ParseVerticalAlign(value);
                break;
            case "text-align":
                declaration.TextAlignment = ParseTextAlign(value);
                break;
            case "color":
                declaration.ForegroundColor = ParseColor(value);
                break;
            case "background":
            case "background-color":
                declaration.BackgroundColor = ParseColor(value);
                break;
        }
    }

    private static IEnumerable<string> SplitDeclarations(string style) {
        int start = 0;
        char quote = '\0';
        int parentheses = 0;
        for (int index = 0; index < style.Length; index++) {
            char current = style[index];
            if (quote != '\0') {
                if (current == quote) {
                    quote = '\0';
                }

                continue;
            }

            if (current == '"' || current == '\'') {
                quote = current;
            } else if (current == '(') {
                parentheses++;
            } else if (current == ')' && parentheses > 0) {
                parentheses--;
            } else if (current == ';' && parentheses == 0) {
                yield return style.Substring(start, index - start);
                start = index + 1;
            }
        }

        if (start < style.Length) {
            yield return style.Substring(start);
        }
    }

    private static string NormalizeValue(string value) {
        string normalized = value.Trim().ToLowerInvariant();
        int important = normalized.IndexOf("!important", StringComparison.OrdinalIgnoreCase);
        return important < 0 ? normalized : normalized.Substring(0, important).TrimEnd();
    }

    private static bool? ParseFontWeight(string value) {
        if (value == "bold" || value == "bolder") {
            return true;
        }

        if (value == "normal" || value == "lighter") {
            return false;
        }

        if (int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int weight)) {
            return weight >= 600;
        }

        return null;
    }

    private static bool? ParseFontStyle(string value) {
        if (value == "italic" || value == "oblique") {
            return true;
        }

        if (value == "normal") {
            return false;
        }

        return null;
    }

    private static void ApplyTextDecoration(HtmlStyleDeclaration declaration, string value) {
        string normalized = " " + value.Replace('-', ' ') + " ";
        if (ContainsWord(normalized, "none")) {
            declaration.Underline = false;
            declaration.Strike = false;
            return;
        }

        if (ContainsWord(normalized, "underline")) {
            declaration.Underline = true;
        }

        if (ContainsWord(normalized, "line through")) {
            declaration.Strike = true;
        }
    }

    private static RtfVerticalPosition? ParseVerticalAlign(string value) {
        switch (value) {
            case "super":
            case "text-top":
                return RtfVerticalPosition.Superscript;
            case "sub":
            case "text-bottom":
                return RtfVerticalPosition.Subscript;
            case "baseline":
            case "middle":
                return RtfVerticalPosition.Baseline;
            default:
                return null;
        }
    }

    private static RtfTextAlignment? ParseTextAlign(string value) {
        switch (value) {
            case "left":
            case "start":
                return RtfTextAlignment.Left;
            case "center":
                return RtfTextAlignment.Center;
            case "right":
            case "end":
                return RtfTextAlignment.Right;
            case "justify":
                return RtfTextAlignment.Justify;
            default:
                return null;
        }
    }

    private static RtfColor? ParseColor(string value) {
        if (value == "transparent" || value == "inherit" || value == "initial" || value == "currentcolor") {
            return null;
        }

        bool isRgbFunction = value.StartsWith("rgb(", StringComparison.Ordinal) || value.StartsWith("rgba(", StringComparison.Ordinal);
        int whitespace = value.IndexOfAny(new[] { ' ', '\t', '\r', '\n' });
        string token = whitespace > 0 && !isRgbFunction ? value.Substring(0, whitespace) : value;
        if (TryParseHexColor(token, out RtfColor? hexColor)) {
            return hexColor;
        }

        if (TryParseRgbColor(value, out RtfColor? rgbColor)) {
            return rgbColor;
        }

        return TryParseNamedColor(token, out RtfColor? namedColor) ? namedColor : null;
    }

    private static bool TryParseHexColor(string value, out RtfColor? color) {
        color = null;
        if (!value.StartsWith("#", StringComparison.Ordinal) || (value.Length != 4 && value.Length != 7)) {
            return false;
        }

        if (value.Length == 4) {
            if (!TryParseHexNibble(value[1], out byte r) ||
                !TryParseHexNibble(value[2], out byte g) ||
                !TryParseHexNibble(value[3], out byte b)) {
                return false;
            }

            color = new RtfColor((byte)(r * 17), (byte)(g * 17), (byte)(b * 17));
            return true;
        }

        if (!byte.TryParse(value.Substring(1, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out byte red) ||
            !byte.TryParse(value.Substring(3, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out byte green) ||
            !byte.TryParse(value.Substring(5, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out byte blue)) {
            return false;
        }

        color = new RtfColor(red, green, blue);
        return true;
    }

    private static bool TryParseRgbColor(string value, out RtfColor? color) {
        color = null;
        int open = value.IndexOf('(');
        int close = value.LastIndexOf(')');
        if (open < 0 || close <= open) {
            return false;
        }

        string function = value.Substring(0, open).Trim();
        if (function != "rgb" && function != "rgba") {
            return false;
        }

        string[] parts = value.Substring(open + 1, close - open - 1).Split(',');
        if (parts.Length < 3 ||
            !TryParseCssByte(parts[0], out byte red) ||
            !TryParseCssByte(parts[1], out byte green) ||
            !TryParseCssByte(parts[2], out byte blue)) {
            return false;
        }

        color = new RtfColor(red, green, blue);
        return true;
    }

    private static bool TryParseCssByte(string value, out byte component) {
        string normalized = value.Trim();
        if (normalized.EndsWith("%", StringComparison.Ordinal)) {
            if (double.TryParse(normalized.Substring(0, normalized.Length - 1), NumberStyles.Float, CultureInfo.InvariantCulture, out double percent)) {
                component = ClampByte((int)Math.Round(percent * 2.55d, MidpointRounding.AwayFromZero));
                return true;
            }

            component = 0;
            return false;
        }

        if (double.TryParse(normalized, NumberStyles.Float, CultureInfo.InvariantCulture, out double number)) {
            component = ClampByte((int)Math.Round(number, MidpointRounding.AwayFromZero));
            return true;
        }

        component = 0;
        return false;
    }

    private static bool TryParseNamedColor(string value, out RtfColor? color) {
        switch (value) {
            case "black":
                color = new RtfColor(0, 0, 0);
                return true;
            case "white":
                color = new RtfColor(255, 255, 255);
                return true;
            case "red":
                color = new RtfColor(255, 0, 0);
                return true;
            case "green":
                color = new RtfColor(0, 128, 0);
                return true;
            case "blue":
                color = new RtfColor(0, 0, 255);
                return true;
            case "yellow":
                color = new RtfColor(255, 255, 0);
                return true;
            case "orange":
                color = new RtfColor(255, 165, 0);
                return true;
            case "purple":
                color = new RtfColor(128, 0, 128);
                return true;
            case "gray":
            case "grey":
                color = new RtfColor(128, 128, 128);
                return true;
            default:
                color = null;
                return false;
        }
    }

    private static bool TryParseHexNibble(char value, out byte result) {
        if (value >= '0' && value <= '9') {
            result = (byte)(value - '0');
            return true;
        }

        if (value >= 'a' && value <= 'f') {
            result = (byte)(value - 'a' + 10);
            return true;
        }

        if (value >= 'A' && value <= 'F') {
            result = (byte)(value - 'A' + 10);
            return true;
        }

        result = 0;
        return false;
    }

    private static byte ClampByte(int value) {
        return (byte)Math.Max(0, Math.Min(255, value));
    }

    private static bool ContainsWord(string value, string word) {
        return value.IndexOf(" " + word + " ", StringComparison.Ordinal) >= 0;
    }
}
