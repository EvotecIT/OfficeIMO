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

    private static bool ContainsWord(string value, string word) {
        return value.IndexOf(" " + word + " ", StringComparison.Ordinal) >= 0;
    }
}
