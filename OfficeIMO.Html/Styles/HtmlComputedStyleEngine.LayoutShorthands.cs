namespace OfficeIMO.Html;

public static partial class HtmlComputedStyleEngine {
    private const string LayoutDeclarationSentinelPrefix = "-officeimo-internal-layout-declaration-";
    private static readonly HashSet<string> LayoutDeclarationNames = new(StringComparer.Ordinal) {
        "grid-column", "grid-row", "grid-area", "grid-column-start", "grid-column-end", "grid-row-start", "grid-row-end",
        "grid-template-columns", "grid-template-rows", "grid-template-areas", "gap", "row-gap", "column-gap",
        "place-items", "place-self", "place-content", "align-items", "justify-items", "align-self", "justify-self", "align-content", "justify-content",
        "text-decoration", "text-decoration-line", "text-decoration-style", "text-decoration-color", "text-decoration-thickness", "text-underline-offset", "text-underline-position", "text-decoration-skip-ink"
    };
    private static readonly string[] GapLonghands = { "row-gap", "column-gap" };

    private static string[]? GetDeferredLayoutShorthandLonghands(string propertyName) => propertyName.ToLowerInvariant() switch {
        "margin" => MarginLonghands,
        "padding" => PaddingLonghands,
        "border" => BorderWidthLonghands.Concat(BorderStyleLonghands).Concat(BorderColorLonghands).ToArray(),
        "border-width" => BorderWidthLonghands,
        "border-style" => BorderStyleLonghands,
        "border-color" => BorderColorLonghands,
        "border-top" => new[] { "border-top-width", "border-top-style", "border-top-color" },
        "border-right" => new[] { "border-right-width", "border-right-style", "border-right-color" },
        "border-bottom" => new[] { "border-bottom-width", "border-bottom-style", "border-bottom-color" },
        "border-left" => new[] { "border-left-width", "border-left-style", "border-left-color" },
        "gap" => GapLonghands,
        "grid-column" => GridColumnLonghands,
        "grid-row" => GridRowLonghands,
        "grid-area" => GridAreaLonghands,
        "place-items" => PlaceItemsLonghands,
        "place-self" => PlaceSelfLonghands,
        "place-content" => PlaceContentLonghands,
        "text-decoration" => TextDecorationLonghands,
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

    private static bool IsGapComponentSyntax(string value) {
        string trimmed = value.Trim();
        if (string.Equals(trimmed, "normal", StringComparison.OrdinalIgnoreCase)) return true;
        // Math syntax is validated here; its sign depends on the real font and
        // containing block. Nonnegative used values are resolved by layout.
        bool unitlessZero = double.TryParse(trimmed, System.Globalization.NumberStyles.Float,
            System.Globalization.CultureInfo.InvariantCulture, out double numeric) && numeric == 0D;
        return (unitlessZero || HtmlRenderCssValues.HasExplicitLengthSyntax(trimmed, allowPercentage: true, allowUnitlessZero: true))
            && TryValidateCssLength(trimmed, out double length)
            && (trimmed.IndexOf('(') >= 0 || length >= 0D);
    }

    private static void ResolveDeferredLayoutLonghands(Dictionary<string, string> properties, IReadOnlyDictionary<string, string> deferred,
        IReadOnlyDictionary<string, string>? parentProperties, ISet<string> inherited, ISet<string> reset, ISet<string> specified,
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
                if (resolved.InheritsComputedValue) {
                    inherited.Add(pending.Key);
                    specified.Remove(pending.Key);
                } else {
                    inherited.Remove(pending.Key);
                    specified.Add(pending.Key);
                }
            } else {
                properties.Remove(pending.Key);
                inherited.Remove(pending.Key);
                specified.Remove(pending.Key);
                reset.Add(pending.Key);
            }
        }
    }

}
