namespace OfficeIMO.Html;

public static partial class HtmlComputedStyleEngine {
    private const string LayoutDeclarationSentinelPrefix = "-officeimo-internal-layout-declaration-";
    private static readonly HashSet<string> LayoutDeclarationNames = new(StringComparer.Ordinal) {
        "grid-column", "grid-row", "grid-area", "grid-column-start", "grid-column-end", "grid-row-start", "grid-row-end",
        "grid-template-columns", "grid-template-rows", "grid-template-areas", "gap", "row-gap", "column-gap"
    };
    private static readonly string[] GapLonghands = { "row-gap", "column-gap" };

    private static string[]? GetDeferredLayoutShorthandLonghands(string propertyName) => propertyName.ToLowerInvariant() switch {
        "gap" => GapLonghands,
        "grid-column" => GridColumnLonghands,
        "grid-row" => GridRowLonghands,
        "grid-area" => GridAreaLonghands,
        _ => null
    };

    private static bool TryExpandGapShorthand(string value, out IReadOnlyList<KeyValuePair<string, string>> longhands) {
        string trimmed = value.Trim();
        if (IsCssWideKeyword(trimmed)) {
            longhands = GapLonghands.Select(name => new KeyValuePair<string, string>(name, trimmed)).ToArray();
            return true;
        }
        IReadOnlyList<string> parts = HtmlRenderCssValues.SplitWhitespace(trimmed);
        if (parts.Count is < 1 or > 2 || parts.Any(part => !IsGapComponentSyntax(part))) {
            longhands = Array.Empty<KeyValuePair<string, string>>();
            return false;
        }
        longhands = new[] {
            new KeyValuePair<string, string>("row-gap", parts[0]),
            new KeyValuePair<string, string>("column-gap", parts.Count == 2 ? parts[1] : parts[0])
        };
        return true;
    }

    private static bool IsGapComponentSyntax(string value) =>
        string.Equals(value.Trim(), "normal", StringComparison.OrdinalIgnoreCase)
        || IsNonNegativeCssLengthOrPercentage(value.Trim());

    private static void ResolveDeferredLayoutLonghands(Dictionary<string, string> properties, IReadOnlyDictionary<string, string> deferred,
        IReadOnlyDictionary<string, string>? parentProperties, ISet<string> inherited, ISet<string> reset,
        bool enforceResolutionLimits) {
        foreach (KeyValuePair<string, string> pending in deferred) {
            string value = "unset";
            if (HtmlCssCustomPropertyResolver.TryResolve(properties[pending.Key],
                    customName => properties.TryGetValue(customName, out string? local) ? local
                        : parentProperties != null && parentProperties.TryGetValue(customName, out string? parent) ? parent : null,
                    out string shorthand, enforceResolutionLimits)
                && TryExpandCascadeShorthand(pending.Value, shorthand, out IReadOnlyList<KeyValuePair<string, string>> longhands)) {
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

}
