namespace OfficeIMO.Html;

public static partial class HtmlComputedStyleEngine {
    private static readonly string[] GridColumnLonghands = { "grid-column-start", "grid-column-end" };
    private static readonly string[] GridRowLonghands = { "grid-row-start", "grid-row-end" };
    private static readonly string[] GridAreaLonghands = { "grid-row-start", "grid-column-start", "grid-row-end", "grid-column-end" };

    private static bool TryExpandGridShorthand(string propertyName, string value, out IReadOnlyList<KeyValuePair<string, string>> longhands) {
        string[] names = GetDeferredLayoutShorthandLonghands(propertyName)!;
        if (IsCssWideKeyword(value.Trim())) {
            longhands = names.Select(name => new KeyValuePair<string, string>(name, value)).ToArray();
            return true;
        }
        IReadOnlyList<string> parts = HtmlRenderCssValues.SplitTopLevel(value, '/');
        if (parts.Count == 0 || parts.Count > names.Length || parts.Any(part => !IsGridLineSyntax(part.Trim()))) {
            longhands = Array.Empty<KeyValuePair<string, string>>();
            return false;
        }
        var values = new string[names.Length];
        values[0] = parts[0].Trim();
        for (int index = 1; index < values.Length; index++) {
            int defaultIndex = propertyName == "grid-area" && index == 3 ? 1 : 0;
            values[index] = parts.Count > index ? parts[index].Trim()
                : IsGridCustomIdentifier(values[defaultIndex]) ? values[defaultIndex] : "auto";
        }
        longhands = names.Select((name, index) => new KeyValuePair<string, string>(name, values[index])).ToArray();
        return true;
    }

    private static bool IsGridCustomIdentifier(string value) {
        if (string.Equals(value, "auto", StringComparison.OrdinalIgnoreCase) || string.Equals(value, "span", StringComparison.OrdinalIgnoreCase)
            || string.Equals(value, "default", StringComparison.OrdinalIgnoreCase) || IsCssWideKeyword(value)) return false;
        int position = 0;
        return HtmlCssIdentifierParser.TryRead(value, ref position, out _) && position == value.Length;
    }

    private static bool IsGridLineSyntax(string value) {
        IReadOnlyList<string> tokens = HtmlRenderCssValues.SplitWhitespace(value);
        if (tokens.Count == 1 && string.Equals(tokens[0], "auto", StringComparison.OrdinalIgnoreCase)) return true;
        if (tokens.Count == 0 || tokens.Count > 3) return false;
        bool span = false;
        bool number = false;
        bool identifier = false;
        long integer = 0;
        foreach (string token in tokens) {
            if (string.Equals(token, "span", StringComparison.OrdinalIgnoreCase)) {
                if (span) return false;
                span = true;
            } else if (long.TryParse(token, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out long parsed)) {
                if (number || parsed == 0) return false;
                number = true;
                integer = parsed;
            } else if (IsGridCustomIdentifier(token)) {
                if (identifier) return false;
                identifier = true;
            } else {
                return false;
            }
        }
        return (number || identifier) && (!span || !number || integer > 0);
    }
}
