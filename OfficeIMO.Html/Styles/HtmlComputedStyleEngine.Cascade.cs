namespace OfficeIMO.Html;

public static partial class HtmlComputedStyleEngine {
    private static readonly OfficeIMO.Html.Css.HtmlCssTokenizationOptions UnboundedPropertyTokenization =
        new OfficeIMO.Html.Css.HtmlCssTokenizationOptions { MaxInputCharacters = null, MaxTokens = null };

    private static void ApplyInlineDeclarations(
        IDictionary<string, CascadedProperty> properties,
        IReadOnlyDictionary<string, string>? parentProperties,
        string? styleText,
        IReadOnlyDictionary<string, CustomPropertyRegistration>? customPropertyRegistrations,
        HtmlCssProcessingBudget budget) {
        if (string.IsNullOrWhiteSpace(styleText)) {
            return;
        }

        OfficeIMO.Html.Css.HtmlCssStyleBlock block;
        try {
            block = OfficeIMO.Html.Css.HtmlCssSyntaxParser.ParseStyleBlock(styleText!, budget.CreateInlineSyntaxOptions());
        } catch (OfficeIMO.Html.Css.HtmlCssSyntaxLimitException exception) {
            throw budget.TranslateSyntaxLimit(exception);
        }
        budget.RecordDeclarations(block.Declarations.Count);
        int declarationOrder = 0;
        foreach (OfficeIMO.Html.Css.HtmlCssDeclaration declaration in block.Declarations) {
            OfficeIMO.Html.Css.HtmlCssPropertyParseResult parsed =
                OfficeIMO.Html.Css.HtmlCssPropertyParser.Parse(declaration, UnboundedPropertyTokenization);
            string value = StripCssCommentsOutsideStrings(parsed.AuthoredValue).Trim();
            if (declaration.Name.Length > 0 && value.Length > 0) {
                ApplyDeclaration(properties, parentProperties, declaration.Name, value, declaration.IsImportant,
                    Specificity.Inline, int.MaxValue, layerOrder: null, declarationOrder: declarationOrder,
                    customPropertyRegistrations: customPropertyRegistrations,
                    source: OfficeIMO.Html.Css.HtmlCssCascadeSourceKind.InlineStyle);
            }
            declarationOrder++;
        }
    }

    private static void ApplyDeclaration(IDictionary<string, CascadedProperty> properties, IReadOnlyDictionary<string, string>? parentProperties,
        string name, string value, bool isImportant, Specificity specificity, int order, CascadeLayerOrder? layerOrder,
        bool valueAlreadyValidated = false, int declarationOrder = 0,
        IReadOnlyDictionary<string, CustomPropertyRegistration>? customPropertyRegistrations = null,
        bool deferredFontShorthand = false,
        OfficeIMO.Html.Css.HtmlCssCascadeSourceKind source = OfficeIMO.Html.Css.HtmlCssCascadeSourceKind.StyleRule,
        string? selector = null,
        string? layerName = null) {
        if (string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(value)) {
            return;
        }

        if (string.Equals(name, "font", StringComparison.OrdinalIgnoreCase)
            && HtmlCssCustomPropertyResolver.ContainsVarFunction(value)) {
            foreach (string longhand in FontShorthandLonghands) {
                ApplyDeclaration(properties, parentProperties, longhand, value, isImportant, specificity, order, layerOrder,
                    valueAlreadyValidated: true, declarationOrder: declarationOrder,
                    customPropertyRegistrations: customPropertyRegistrations, deferredFontShorthand: true,
                    source: source, selector: selector, layerName: layerName);
            }
        }

        if (!HtmlCssCustomPropertyResolver.ContainsVarFunction(value)
            && (valueAlreadyValidated || IsSupportedDeclarationValue(name, value))
            && TryExpandCascadeShorthand(name, value, out IReadOnlyList<KeyValuePair<string, string>> boxLonghands)) {
            foreach (KeyValuePair<string, string> longhand in boxLonghands) {
                ApplyDeclaration(properties, parentProperties, longhand.Key, longhand.Value, isImportant, specificity, order, layerOrder,
                    valueAlreadyValidated: false, declarationOrder: declarationOrder, customPropertyRegistrations: customPropertyRegistrations,
                    source: source, selector: selector, layerName: layerName);
            }
        }

        string shorthandValue = value;
        if (string.Equals(name, "container", StringComparison.OrdinalIgnoreCase)
            && HtmlCssCustomPropertyResolver.ContainsVarFunction(value)) {
            HtmlCssCustomPropertyResolver.TryResolve(
                value,
                customName => TryGetCascadedValue(properties, customName)
                    ?? (parentProperties != null && parentProperties.TryGetValue(customName, out string? inherited) ? inherited : null),
                out shorthandValue);
        }
        if (string.Equals(name, "container", StringComparison.OrdinalIgnoreCase)
            && IsSupportedDeclarationValue(name, shorthandValue)
            && TryExpandContainerShorthand(shorthandValue, out string containerName, out string containerType)) {
            ApplyDeclaration(properties, parentProperties, "container-name", containerName, isImportant, specificity, order, layerOrder, declarationOrder: declarationOrder, customPropertyRegistrations: customPropertyRegistrations, source: source, selector: selector, layerName: layerName);
            ApplyDeclaration(properties, parentProperties, "container-type", containerType, isImportant, specificity, order, layerOrder, declarationOrder: declarationOrder, customPropertyRegistrations: customPropertyRegistrations, source: source, selector: selector, layerName: layerName);
        }
        if (string.Equals(name, "animation", StringComparison.OrdinalIgnoreCase)
            && IsSupportedDeclarationValue(name, value)
            && HtmlResourcePipeline.TryExpandAnimationShorthandNames(value, out string animationNames)) {
            ApplyDeclaration(properties, parentProperties, "animation-name", animationNames, isImportant, specificity, order, layerOrder, declarationOrder: declarationOrder, customPropertyRegistrations: customPropertyRegistrations, source: source, selector: selector, layerName: layerName);
        }
        string imageSourceProperty = GetImageSourcePropertyName(name);
        if (!string.Equals(imageSourceProperty, name, StringComparison.OrdinalIgnoreCase)
            && IsSupportedDeclarationValue(name, value)) {
            ApplyDeclaration(properties, parentProperties, imageSourceProperty, value, isImportant, specificity, order, layerOrder, declarationOrder: declarationOrder, customPropertyRegistrations: customPropertyRegistrations, source: source, selector: selector, layerName: layerName);
        }

        CascadedProperty? existing;
        properties.TryGetValue(name, out existing);
        if (string.Equals(value.Trim(), "revert-layer", StringComparison.OrdinalIgnoreCase)) {
            var candidate = CascadedProperty.RevertLayer(isImportant, specificity, order, layerOrder, alternatives: null,
                declarationOrder, value, source, selector, layerName);
            if (existing != null && !ShouldReplace(existing, isImportant, specificity, order, layerOrder, declarationOrder)) {
                properties[name] = existing.WithAlternative(candidate);
                return;
            }
            properties[name] = CascadedProperty.RevertLayer(isImportant, specificity, order, layerOrder, CollectCandidates(existing),
                declarationOrder, value, source, selector, layerName);
            return;
        }

        var resolved = ResolveCssWideKeyword(name, value, parentProperties, customPropertyRegistrations);
        if (!resolved.HasValue) {
            CascadedProperty? resetExisting;
            if (properties.TryGetValue(name, out resetExisting) && resetExisting != null && !ShouldReplace(resetExisting, isImportant, specificity, order, layerOrder, declarationOrder)) {
                resetExisting = resetExisting.WithAlternative(CascadedProperty.Clear(isImportant, specificity, order, layerOrder,
                    alternatives: null, declarationOrder, value, source, selector, layerName));
                properties[name] = resetExisting;
                return;
            }

            properties[name] = CascadedProperty.Clear(isImportant, specificity, order, layerOrder, CollectCandidates(resetExisting),
                declarationOrder, value, source, selector, layerName);
            return;
        }

        if (!valueAlreadyValidated && !IsSupportedDeclarationValue(name, resolved.Value)) {
            return;
        }

        if (existing != null && !ShouldReplace(existing, isImportant, specificity, order, layerOrder, declarationOrder)) {
            properties[name] = existing.WithAlternative(new CascadedProperty(resolved.Value, isImportant, specificity, order, layerOrder,
                inheritsComputedValue: resolved.InheritsComputedValue, declarationOrder: declarationOrder,
                deferredFontShorthand: deferredFontShorthand, authoredValue: value, source: source, selector: selector, layerName: layerName));
            return;
        }

        properties[name] = new CascadedProperty(resolved.Value, isImportant, specificity, order, layerOrder, CollectCandidates(existing),
            resolved.InheritsComputedValue, declarationOrder, deferredFontShorthand, value, source, selector, layerName);
    }

    private static string? TryGetCascadedValue(IDictionary<string, CascadedProperty> properties, string name) {
        if (!properties.TryGetValue(name, out CascadedProperty? property)) return null;
        return ResolveLayerRevert(property)?.HasValue == true ? ResolveLayerRevert(property)!.Value : null;
    }

    internal static string GetImageSourcePropertyName(string propertyName) {
        if (string.Equals(propertyName, "background", StringComparison.OrdinalIgnoreCase)) return "background-image";
        return propertyName;
    }

    private static bool TryExpandContainerShorthand(string value, out string containerName, out string containerType) {
        string normalized = value.Trim();
        if (IsCssWideKeyword(normalized)) {
            containerName = normalized;
            containerType = normalized;
            return true;
        }
        if (HtmlCssCustomPropertyResolver.ContainsVarFunction(normalized)) {
            containerName = string.Empty;
            containerType = string.Empty;
            return false;
        }
        int slash = normalized.IndexOf('/');
        containerName = (slash < 0 ? normalized : normalized.Substring(0, slash)).Trim();
        containerType = slash < 0 ? "normal" : normalized.Substring(slash + 1).Trim();
        return containerName.Length > 0 && containerType.Length > 0;
    }

    private static CssKeywordResolution ResolveCssWideKeyword(
        string name,
        string value,
        IReadOnlyDictionary<string, string>? parentProperties,
        IReadOnlyDictionary<string, CustomPropertyRegistration>? customPropertyRegistrations = null) {
        string trimmed = value.Trim();
        if (string.Equals(trimmed, "inherit", StringComparison.OrdinalIgnoreCase)
            || (string.Equals(trimmed, "unset", StringComparison.OrdinalIgnoreCase) && IsInheritedProperty(name, customPropertyRegistrations))) {
            string? inheritedValue;
            return parentProperties != null && parentProperties.TryGetValue(name, out inheritedValue) && !string.IsNullOrWhiteSpace(inheritedValue)
                ? CssKeywordResolution.ForInheritedValue(inheritedValue)
                : CssKeywordResolution.Clear;
        }

        if (string.Equals(trimmed, "revert", StringComparison.OrdinalIgnoreCase) && IsInheritedProperty(name, customPropertyRegistrations)) {
            string? inheritedValue;
            return parentProperties != null && parentProperties.TryGetValue(name, out inheritedValue) && !string.IsNullOrWhiteSpace(inheritedValue)
                ? CssKeywordResolution.ForInheritedValue(inheritedValue)
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

    private static bool ShouldReplace(CascadedProperty existing, bool isImportant, Specificity specificity, int order, CascadeLayerOrder? layerOrder, int declarationOrder = 0) {
        // Inheritance happens after the cascade. A value copied from the parent is therefore
        // only a fallback for this element and must never outrank a declaration that matches
        // the element, including a declaration inside a cascade layer.
        if (ReferenceEquals(existing.Specificity, Specificity.Inherited)) {
            return true;
        }

        if (existing.IsImportant != isImportant) {
            return isImportant;
        }

        if (isImportant) {
            bool existingInline = ReferenceEquals(existing.Specificity, Specificity.Inline);
            bool candidateInline = ReferenceEquals(specificity, Specificity.Inline);
            if (existingInline != candidateInline) return candidateInline;
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

        if (order != existing.Order) return order > existing.Order;
        return declarationOrder >= existing.DeclarationOrder;
    }

    private static IReadOnlyList<CascadedProperty> CollectCandidates(CascadedProperty? property) {
        if (property == null) return Array.Empty<CascadedProperty>();
        var candidates = new List<CascadedProperty>(property.Alternatives.Count + 1) { property };
        candidates.AddRange(property.Alternatives);
        return candidates;
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
