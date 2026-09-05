using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using System.Globalization;

namespace OfficeIMO.Html;

public static partial class HtmlComputedStyleEngine {
    private static IReadOnlyDictionary<string, CustomPropertyRegistration> ParseCustomPropertyRegistrations(
        IHtmlDocument document,
        MediaEnvironment environment) {
        var registrations = new Dictionary<string, CustomPropertyRegistration>(HtmlCssPropertyNameComparer.Instance);
        foreach (IElement styleElement in document.QuerySelectorAll("style")) {
            if (!IsCssStyleElement(styleElement)
                || !IsApplicableMedia(styleElement.GetAttribute("media") ?? string.Empty, environment)) {
                continue;
            }

            ParseTopLevelCustomPropertyRegistrations(styleElement.TextContent ?? string.Empty, registrations);
        }
        return registrations;
    }

    private static void ParseTopLevelCustomPropertyRegistrations(
        string css,
        IDictionary<string, CustomPropertyRegistration> registrations) {
        int depth = 0;
        char quote = '\0';
        for (int index = 0; index < css.Length; index++) {
            char current = css[index];
            if (quote != '\0') {
                if (current == '\\') index++;
                else if (current == quote) quote = '\0';
                continue;
            }
            if (current == '\'' || current == '"') {
                quote = current;
                continue;
            }
            if (current == '/' && index + 1 < css.Length && css[index + 1] == '*') {
                int commentEnd = css.IndexOf("*/", index + 2, StringComparison.Ordinal);
                if (commentEnd < 0) return;
                index = commentEnd + 1;
                continue;
            }
            if (current == '{') {
                depth++;
                continue;
            }
            if (current == '}') {
                if (depth > 0) depth--;
                continue;
            }
            if (depth != 0 || current != '@' || !MatchesAtRuleName(css, index + 1, "property")) continue;

            int cursor = index + "@property".Length;
            SkipCssWhitespace(css, ref cursor);
            int nameStart = cursor;
            while (cursor < css.Length && css[cursor] != '{' && css[cursor] != ';') cursor++;
            if (cursor >= css.Length || css[cursor] != '{') continue;
            string nameText = css.Substring(nameStart, cursor - nameStart).Trim();
            int close = FindCustomPropertyBlockEnd(css, cursor);
            if (close < 0) return;
            string block = css.Substring(cursor + 1, close - cursor - 1);
            if (TryCreateCustomPropertyRegistration(nameText, block, out CustomPropertyRegistration? registration)) {
                registrations[registration!.Name] = registration;
            }
            index = close;
        }
    }

    private static bool TryCreateCustomPropertyRegistration(
        string nameText,
        string block,
        out CustomPropertyRegistration? registration) {
        registration = null;
        if (!HtmlCssIdentifierParser.TryParse(nameText, out string name)
            || !name.StartsWith("--", StringComparison.Ordinal)
            || name.Length <= 2) {
            return false;
        }

        string? syntax = null;
        bool? inherits = null;
        string? initialValue = null;
        foreach (string declaration in SplitCssDeclarations(StripCssCommentsOutsideStrings(block))) {
            int separator = declaration.IndexOf(':');
            if (separator <= 0) continue;
            string descriptor = declaration.Substring(0, separator).Trim();
            string value = declaration.Substring(separator + 1).Trim();
            if (string.Equals(descriptor, "syntax", StringComparison.OrdinalIgnoreCase)) {
                if (value.Length < 2 || value[0] != value[value.Length - 1] || value[0] != '\'' && value[0] != '"') return false;
                syntax = HtmlCssEscapeDecoder.Decode(value.Substring(1, value.Length - 2)).Trim();
            } else if (string.Equals(descriptor, "inherits", StringComparison.OrdinalIgnoreCase)) {
                if (string.Equals(value, "true", StringComparison.OrdinalIgnoreCase)) inherits = true;
                else if (string.Equals(value, "false", StringComparison.OrdinalIgnoreCase)) inherits = false;
                else return false;
            } else if (string.Equals(descriptor, "initial-value", StringComparison.OrdinalIgnoreCase)) {
                initialValue = value;
            }
        }

        if (string.IsNullOrWhiteSpace(syntax) || !inherits.HasValue) return false;
        if (syntax != "*" && string.IsNullOrWhiteSpace(initialValue)) return false;
        if (!string.IsNullOrWhiteSpace(initialValue)
            && (HtmlCssCustomPropertyResolver.ContainsVarFunction(initialValue!)
                || !IsRegisteredCustomPropertyValueValid(syntax!, initialValue!))) {
            return false;
        }
        registration = new CustomPropertyRegistration(name, syntax!, inherits.Value, initialValue);
        return true;
    }

    private static bool IsRegisteredCustomPropertyValueValid(string syntax, string value) {
        string normalized = value.Trim();
        if (normalized.Length == 0) return false;
        if (syntax == "*") return true;
        foreach (string alternative in syntax.Split('|')) {
            string component = alternative.Trim();
            if (component.Length == 0) continue;
            bool commaSeparated = component.EndsWith("#", StringComparison.Ordinal);
            bool spaceSeparated = component.EndsWith("+", StringComparison.Ordinal);
            if (commaSeparated || spaceSeparated) component = component.Substring(0, component.Length - 1).TrimEnd();
            IReadOnlyList<string> values = commaSeparated
                ? SplitTopLevelCustomPropertyValues(normalized, ',')
                : spaceSeparated ? HtmlRenderCssValues.SplitWhitespace(normalized) : new[] { normalized };
            if (values.Count > 0 && values.All(item => IsRegisteredCustomPropertyComponentValid(component, item))) return true;
        }
        return false;
    }

    private static bool IsRegisteredCustomPropertyComponentValid(string syntax, string value) {
        string normalized = value.Trim();
        switch (syntax.ToLowerInvariant()) {
            case "<color>":
                return normalized.Equals("currentcolor", StringComparison.OrdinalIgnoreCase)
                    || HtmlRenderCssValues.TryColor(normalized, out _);
            case "<length>":
                return HtmlRenderCssValues.HasExplicitLengthSyntax(normalized, false, true)
                    && TryValidateCssLength(normalized, out _);
            case "<percentage>":
                return normalized.EndsWith("%", StringComparison.Ordinal)
                    && double.TryParse(normalized.Substring(0, normalized.Length - 1), NumberStyles.Float, CultureInfo.InvariantCulture, out _);
            case "<length-percentage>":
                return HtmlRenderCssValues.HasExplicitLengthSyntax(normalized, true, true)
                    && TryValidateCssLength(normalized, out _);
            case "<number>":
                return double.TryParse(normalized, NumberStyles.Float, CultureInfo.InvariantCulture, out _);
            case "<integer>":
                return int.TryParse(normalized, NumberStyles.Integer, CultureInfo.InvariantCulture, out _);
            case "<angle>":
                return EndsWithFiniteNumber(normalized, "deg") || EndsWithFiniteNumber(normalized, "grad")
                    || EndsWithFiniteNumber(normalized, "rad") || EndsWithFiniteNumber(normalized, "turn");
            case "<time>":
                return EndsWithFiniteNumber(normalized, "ms") || EndsWithFiniteNumber(normalized, "s");
            case "<resolution>":
                return EndsWithFiniteNumber(normalized, "dpi") || EndsWithFiniteNumber(normalized, "dpcm")
                    || EndsWithFiniteNumber(normalized, "dppx") || EndsWithFiniteNumber(normalized, "x");
            case "<custom-ident>":
                return HtmlCssIdentifierParser.TryParse(normalized, out string identifier)
                    && !IsCssWideKeyword(identifier) && !string.Equals(identifier, "default", StringComparison.OrdinalIgnoreCase);
            case "<url>":
                return normalized.StartsWith("url(", StringComparison.OrdinalIgnoreCase) && normalized.EndsWith(")", StringComparison.Ordinal);
            case "<image>":
                return normalized.Equals("none", StringComparison.OrdinalIgnoreCase)
                    || normalized.StartsWith("url(", StringComparison.OrdinalIgnoreCase)
                    || normalized.IndexOf("gradient(", StringComparison.OrdinalIgnoreCase) >= 0;
            case "<transform-function>":
            case "<transform-list>":
                return IsSupportedDeclarationValue("transform", normalized);
            default:
                return HtmlCssIdentifierParser.TryParse(syntax, out string literal)
                    && string.Equals(literal, normalized, StringComparison.OrdinalIgnoreCase);
        }
    }

    private static bool EndsWithFiniteNumber(string value, string suffix) {
        if (!value.EndsWith(suffix, StringComparison.OrdinalIgnoreCase) || value.Length == suffix.Length) return false;
        return double.TryParse(value.Substring(0, value.Length - suffix.Length), NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed)
            && !double.IsNaN(parsed) && !double.IsInfinity(parsed);
    }

    private static IReadOnlyList<string> SplitTopLevelCustomPropertyValues(string value, char separator) {
        var values = new List<string>();
        int depth = 0;
        char quote = '\0';
        int start = 0;
        for (int index = 0; index < value.Length; index++) {
            char current = value[index];
            if (quote != '\0') {
                if (current == '\\') index++;
                else if (current == quote) quote = '\0';
            } else if (current == '\'' || current == '"') quote = current;
            else if (current == '(' || current == '[') depth++;
            else if ((current == ')' || current == ']') && depth > 0) depth--;
            else if (current == separator && depth == 0) {
                values.Add(value.Substring(start, index - start).Trim());
                start = index + 1;
            }
        }
        values.Add(value.Substring(start).Trim());
        return values;
    }

    private static bool MatchesAtRuleName(string css, int start, string name) {
        if (start + name.Length > css.Length
            || string.Compare(css, start, name, 0, name.Length, StringComparison.OrdinalIgnoreCase) != 0) return false;
        int end = start + name.Length;
        return end >= css.Length || !char.IsLetterOrDigit(css[end]) && css[end] != '-' && css[end] != '_';
    }

    private static void SkipCssWhitespace(string css, ref int cursor) {
        while (cursor < css.Length && char.IsWhiteSpace(css[cursor])) cursor++;
    }

    private static int FindCustomPropertyBlockEnd(string css, int open) {
        int depth = 1;
        char quote = '\0';
        for (int index = open + 1; index < css.Length; index++) {
            char current = css[index];
            if (quote != '\0') {
                if (current == '\\') index++;
                else if (current == quote) quote = '\0';
                continue;
            }
            if (current == '\'' || current == '"') quote = current;
            else if (current == '/' && index + 1 < css.Length && css[index + 1] == '*') {
                int commentEnd = css.IndexOf("*/", index + 2, StringComparison.Ordinal);
                if (commentEnd < 0) return -1;
                index = commentEnd + 1;
            } else if (current == '{') depth++;
            else if (current == '}' && --depth == 0) return index;
        }
        return -1;
    }
}
