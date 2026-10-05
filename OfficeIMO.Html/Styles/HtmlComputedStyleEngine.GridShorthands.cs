namespace OfficeIMO.Html;

public static partial class HtmlComputedStyleEngine {
    private const string GridDeclarationSentinelPrefix = "-officeimo-internal-grid-declaration-";
    private static readonly HashSet<string> GridDeclarationNames = new(StringComparer.Ordinal) {
        "grid-column", "grid-row", "grid-area", "grid-column-start", "grid-column-end", "grid-row-start", "grid-row-end",
        "grid-template-columns", "grid-template-rows", "grid-template-areas"
    };

    private static readonly string[] GridColumnLonghands = { "grid-column-start", "grid-column-end" };
    private static readonly string[] GridRowLonghands = { "grid-row-start", "grid-row-end" };
    private static readonly string[] GridAreaLonghands = { "grid-row-start", "grid-column-start", "grid-row-end", "grid-column-end" };

    private static string[]? GetGridShorthandLonghands(string propertyName) => propertyName.ToLowerInvariant() switch {
        "grid-column" => GridColumnLonghands,
        "grid-row" => GridRowLonghands,
        "grid-area" => GridAreaLonghands,
        _ => null
    };

    private static bool TryExpandGridShorthand(string propertyName, string value, out IReadOnlyList<KeyValuePair<string, string>> longhands) {
        string[] names = GetGridShorthandLonghands(propertyName)!;
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

    private static void ResolveDeferredGridLonghands(Dictionary<string, string> properties, IReadOnlyDictionary<string, string> deferred,
        IReadOnlyDictionary<string, string>? parentProperties, ISet<string> inherited, ISet<string> reset,
        bool enforceResolutionLimits) {
        foreach (KeyValuePair<string, string> pending in deferred) {
            string value = "unset";
            if (HtmlCssCustomPropertyResolver.TryResolve(properties[pending.Key],
                    customName => properties.TryGetValue(customName, out string? local) ? local
                        : parentProperties != null && parentProperties.TryGetValue(customName, out string? parent) ? parent : null,
                    out string shorthand, enforceResolutionLimits)
                && TryExpandGridShorthand(pending.Value, shorthand, out IReadOnlyList<KeyValuePair<string, string>> longhands)) {
                value = longhands.First(item => item.Key == pending.Key).Value;
            }
            // Invalid substitution still occupies the shorthand's cascade position.
            CssKeywordResolution resolved = ResolveCssWideKeyword(pending.Key, value, parentProperties);
            if (resolved.HasValue) {
                properties[pending.Key] = resolved.Value;
                reset.Remove(pending.Key);
                if (resolved.InheritsComputedValue) inherited.Add(pending.Key); else inherited.Remove(pending.Key);
            } else {
                properties.Remove(pending.Key);
                inherited.Remove(pending.Key);
                reset.Add(pending.Key);
            }
        }
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
