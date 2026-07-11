namespace OfficeIMO.Html;

public static partial class HtmlComputedStyleEngine {
    private static void ApplyInlineDeclarations(IDictionary<string, CascadedProperty> properties, IReadOnlyDictionary<string, string>? parentProperties, string? styleText) {
        if (string.IsNullOrWhiteSpace(styleText)) {
            return;
        }

        foreach (string declaration in SplitCssDeclarations(StripCssCommentsOutsideStrings(styleText!))) {
            int separator = declaration.IndexOf(':');
            if (separator <= 0) {
                continue;
            }

            string name = declaration.Substring(0, separator).Trim();
            string value = declaration.Substring(separator + 1).Trim();
            bool isImportant;
            value = StripTrailingImportant(value, out isImportant);

            if (name.Length > 0 && value.Length > 0) {
                ApplyDeclaration(properties, parentProperties, name, value, isImportant, Specificity.Inline, int.MaxValue);
            }
        }
    }

    private static void ApplyDeclaration(IDictionary<string, CascadedProperty> properties, IReadOnlyDictionary<string, string>? parentProperties, string name, string value, bool isImportant, Specificity specificity, int order) {
        if (string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(value)) {
            return;
        }

        var resolved = ResolveCssWideKeyword(name, value, parentProperties);
        if (!resolved.HasValue) {
            CascadedProperty? resetExisting;
            if (properties.TryGetValue(name, out resetExisting) && resetExisting != null && !ShouldReplace(resetExisting, isImportant, specificity, order)) {
                return;
            }

            properties[name] = CascadedProperty.Clear(isImportant, specificity, order);
            return;
        }

        if (!IsSupportedDeclarationValue(name, resolved.Value)) {
            return;
        }

        CascadedProperty? existing;
        if (properties.TryGetValue(name, out existing) && existing != null && !ShouldReplace(existing, isImportant, specificity, order)) {
            return;
        }

        properties[name] = new CascadedProperty(resolved.Value, isImportant, specificity, order);
    }

    private static CssKeywordResolution ResolveCssWideKeyword(string name, string value, IReadOnlyDictionary<string, string>? parentProperties) {
        string trimmed = value.Trim();
        if (string.Equals(trimmed, "inherit", StringComparison.OrdinalIgnoreCase)
            || (string.Equals(trimmed, "unset", StringComparison.OrdinalIgnoreCase) && InheritedProperties.Contains(name))) {
            string? inheritedValue;
            return parentProperties != null && parentProperties.TryGetValue(name, out inheritedValue) && !string.IsNullOrWhiteSpace(inheritedValue)
                ? CssKeywordResolution.ForValue(inheritedValue)
                : CssKeywordResolution.Clear;
        }

        if (string.Equals(trimmed, "initial", StringComparison.OrdinalIgnoreCase)
            || string.Equals(trimmed, "revert", StringComparison.OrdinalIgnoreCase)
            || string.Equals(trimmed, "revert-layer", StringComparison.OrdinalIgnoreCase)) {
            return string.Equals(name, "visibility", StringComparison.OrdinalIgnoreCase)
                ? CssKeywordResolution.ForValue("visible")
                : CssKeywordResolution.Clear;
        }

        if (string.Equals(trimmed, "unset", StringComparison.OrdinalIgnoreCase)) {
            return CssKeywordResolution.Clear;
        }

        return CssKeywordResolution.ForValue(value);
    }

    private static bool ShouldReplace(CascadedProperty existing, bool isImportant, Specificity specificity, int order) {
        if (existing.IsImportant != isImportant) {
            return isImportant;
        }

        int specificityComparison = specificity.CompareTo(existing.Specificity);
        if (specificityComparison != 0) {
            return specificityComparison > 0;
        }

        return order >= existing.Order;
    }

    private static string StripTrailingImportant(string value, out bool isImportant) {
        isImportant = false;
        if (string.IsNullOrWhiteSpace(value)) {
            return value;
        }

        string trimmed = value.TrimEnd();
        const string ImportantKeyword = "important";
        int importantStart = trimmed.Length - ImportantKeyword.Length;
        if (importantStart < 0 || !string.Equals(trimmed.Substring(importantStart), ImportantKeyword, StringComparison.OrdinalIgnoreCase)) {
            return value;
        }

        int bangIndex = importantStart - 1;
        while (bangIndex >= 0 && char.IsWhiteSpace(trimmed[bangIndex])) {
            bangIndex--;
        }

        if (bangIndex < 0 || trimmed[bangIndex] != '!') {
            return value;
        }

        if (IsInsideCssString(trimmed, bangIndex) || IsInsideCssComment(trimmed, bangIndex)) {
            return value;
        }

        isImportant = true;
        return trimmed.Substring(0, bangIndex).TrimEnd();
    }

}
