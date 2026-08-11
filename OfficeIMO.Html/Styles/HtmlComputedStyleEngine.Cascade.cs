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
                ApplyDeclaration(properties, parentProperties, name, value, isImportant, Specificity.Inline, int.MaxValue, layerOrder: null);
            }
        }
    }

    private static void ApplyDeclaration(IDictionary<string, CascadedProperty> properties, IReadOnlyDictionary<string, string>? parentProperties, string name, string value, bool isImportant, Specificity specificity, int order, CascadeLayerOrder? layerOrder) {
        if (string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(value)) {
            return;
        }

        CascadedProperty? existing;
        properties.TryGetValue(name, out existing);
        if (string.Equals(value.Trim(), "revert-layer", StringComparison.OrdinalIgnoreCase)) {
            var candidate = CascadedProperty.RevertLayer(isImportant, specificity, order, layerOrder, alternatives: null);
            if (existing != null && !ShouldReplace(existing, isImportant, specificity, order, layerOrder)) {
                properties[name] = existing.WithAlternative(candidate);
                return;
            }
            properties[name] = CascadedProperty.RevertLayer(isImportant, specificity, order, layerOrder, CollectCandidates(existing));
            return;
        }

        var resolved = ResolveCssWideKeyword(name, value, parentProperties);
        if (!resolved.HasValue) {
            CascadedProperty? resetExisting;
            if (properties.TryGetValue(name, out resetExisting) && resetExisting != null && !ShouldReplace(resetExisting, isImportant, specificity, order, layerOrder)) {
                resetExisting = resetExisting.WithAlternative(CascadedProperty.Clear(isImportant, specificity, order, layerOrder, alternatives: null));
                properties[name] = resetExisting;
                return;
            }

            properties[name] = CascadedProperty.Clear(isImportant, specificity, order, layerOrder, CollectCandidates(resetExisting));
            return;
        }

        if (!IsSupportedDeclarationValue(name, resolved.Value)) {
            return;
        }

        if (existing != null && !ShouldReplace(existing, isImportant, specificity, order, layerOrder)) {
            properties[name] = existing.WithAlternative(new CascadedProperty(resolved.Value, isImportant, specificity, order, layerOrder));
            return;
        }

        properties[name] = new CascadedProperty(resolved.Value, isImportant, specificity, order, layerOrder, CollectCandidates(existing));
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
            || string.Equals(trimmed, "revert", StringComparison.OrdinalIgnoreCase)) {
            return string.Equals(name, "visibility", StringComparison.OrdinalIgnoreCase)
                ? CssKeywordResolution.ForValue("visible")
                : CssKeywordResolution.Clear;
        }

        if (string.Equals(trimmed, "unset", StringComparison.OrdinalIgnoreCase)) {
            return CssKeywordResolution.Clear;
        }

        return CssKeywordResolution.ForValue(value);
    }

    private static bool ShouldReplace(CascadedProperty existing, bool isImportant, Specificity specificity, int order, CascadeLayerOrder? layerOrder) {
        // Inheritance happens after the cascade. A value copied from the parent is therefore
        // only a fallback for this element and must never outrank a declaration that matches
        // the element, including a declaration inside a cascade layer.
        if (ReferenceEquals(existing.Specificity, Specificity.Inherited)) {
            return true;
        }

        if (existing.IsImportant != isImportant) {
            return isImportant;
        }

        if ((existing.LayerOrder != null) != (layerOrder != null)) {
            return isImportant ? layerOrder != null : layerOrder == null;
        }

        if (existing.LayerOrder != null && layerOrder != null) {
            int layerComparison = layerOrder.CompareTo(existing.LayerOrder);
            if (layerComparison != 0) {
            return isImportant
                    ? layerComparison < 0
                    : layerComparison > 0;
            }
        }

        int specificityComparison = specificity.CompareTo(existing.Specificity);
        if (specificityComparison != 0) {
            return specificityComparison > 0;
        }

        return order >= existing.Order;
    }

    private static IReadOnlyList<CascadedProperty> CollectCandidates(CascadedProperty? property) {
        if (property == null) return Array.Empty<CascadedProperty>();
        var candidates = new List<CascadedProperty>(property.Alternatives.Count + 1) { property };
        candidates.AddRange(property.Alternatives);
        return candidates.AsReadOnly();
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
