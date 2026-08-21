using AngleSharp.Dom;
using AngleSharp.Html.Dom;

namespace OfficeIMO.Html;

public static partial class HtmlEditableLayoutProjector {
    private static bool HasSemanticFlowAncestor(IElement element, ISet<IElement> semanticFlowRoots) {
        for (IElement? ancestor = element.ParentElement; ancestor != null; ancestor = ancestor.ParentElement) {
            if (semanticFlowRoots.Contains(ancestor)) return true;
        }
        return false;
    }

    private static bool ContainsSemanticRichContent(
        IElement element,
        IReadOnlyDictionary<IElement, HtmlComputedStyle> styles) {
        if (string.Equals(element.LocalName, "img", StringComparison.OrdinalIgnoreCase)) {
            return HasVisibleRootBorder(element, styles);
        }
        if (IsSemanticRichElement(element)) return true;
        if (HasLanguageMetadata(element)) return true;
        if (HasDistinctRichTextStyle(element, styles)) return true;
        if (HasVisibleRootBorder(element, styles)) return true;
        return element.QuerySelectorAll("*").Any(child => {
            if (string.Equals(child.LocalName, "img", StringComparison.OrdinalIgnoreCase)) {
                return HasVisibleRootBorder(child, styles);
            }
            return IsSemanticRichElement(child)
                || HasLanguageMetadata(child)
                || HasDistinctRichTextStyle(child, styles)
                || HasDistinctStyle(child, styles, RichDescendantVisualStyleProperties);
        });
    }

    private static bool IsSemanticRichElement(IElement element) =>
        SemanticRichElementNames.Contains(element.LocalName)
        || string.Equals(element.NamespaceUri, "http://www.w3.org/1998/Math/MathML",
            StringComparison.OrdinalIgnoreCase);

    private static bool HasLanguageMetadata(IElement element) =>
        !string.IsNullOrWhiteSpace(element.GetAttribute("lang"))
        || !string.IsNullOrWhiteSpace(element.GetAttribute("xml:lang"));

    private static bool ContainsGeneratedPseudoContent(
        IElement element,
        HtmlComputedStyleSet styles) {
        if (HasGeneratedPseudoContent(element, styles)) return true;
        return element.QuerySelectorAll("*").Any(child => HasGeneratedPseudoContent(child, styles));
    }

    private static bool HasGeneratedPseudoContent(IElement element, HtmlComputedStyleSet styles) =>
        HasGeneratedPseudoContent(element, HtmlPseudoElementKind.Before, styles)
        || HasGeneratedPseudoContent(element, HtmlPseudoElementKind.After, styles);

    private static bool HasGeneratedPseudoContent(
        IElement element,
        HtmlPseudoElementKind kind,
        HtmlComputedStyleSet styles) {
        if (!styles.TryGetPseudoStyle(element, kind, out HtmlComputedStyle style)
            || style.GetValue("display").Trim().Equals("none", StringComparison.OrdinalIgnoreCase)) {
            return false;
        }
        string content = style.GetValue("content").Trim();
        return content.Length > 0
            && !content.Equals("normal", StringComparison.OrdinalIgnoreCase)
            && !content.Equals("none", StringComparison.OrdinalIgnoreCase);
    }

    private static bool ContainsBookmarkTarget(IElement element) {
        if (!string.IsNullOrWhiteSpace(element.GetAttribute("id"))) return true;
        return element.QuerySelectorAll("[id]").Any(target =>
            !string.IsNullOrWhiteSpace(target.GetAttribute("id")));
    }

    private static bool ContainsHtmlComment(IElement element) =>
        element.Descendants<IComment>().Any();

    private static bool IsSemanticSectionOwner(IElement element, IHtmlDocument document) {
        if (!element.LocalName.Equals("section", StringComparison.OrdinalIgnoreCase)
            && !element.LocalName.Equals("article", StringComparison.OrdinalIgnoreCase)) {
            return false;
        }
        IElement? parent = element.ParentElement;
        while (parent != null && parent.LocalName.Equals("main", StringComparison.OrdinalIgnoreCase)) {
            parent = parent.ParentElement;
        }
        return ReferenceEquals(parent, document.Body)
            || ReferenceEquals(parent, document.DocumentElement);
    }

    private static bool ContainsMixedInlineImageContent(
        IElement element,
        IReadOnlyDictionary<IElement, HtmlComputedStyle> styles,
        bool preserveEdgeSequences) {
        var contentOrder = new List<bool>();
        AppendInlineContentOrder(element, element, styles, contentOrder);
        if (!contentOrder.Contains(true) || !contentOrder.Contains(false)) return false;
        if (preserveEdgeSequences) return true;
        for (int index = 1; index < contentOrder.Count - 1; index++) {
            if (contentOrder[index]
                && contentOrder.Take(index).Contains(false)
                && contentOrder.Skip(index + 1).Contains(false)) return true;
        }
        return false;
    }

    private static bool HasMultipleVisibleLayoutChildren(
        IElement element,
        IReadOnlyDictionary<IElement, HtmlComputedStyle> styles) {
        int visibleChildren = 0;
        foreach (INode child in element.ChildNodes) {
            if (child is IElement childElement) {
                if (!IsProjectionElementVisible(childElement, element, styles)
                    || (string.IsNullOrWhiteSpace(childElement.TextContent)
                        && childElement.QuerySelector("img") == null)) continue;
            } else if (string.IsNullOrWhiteSpace(child.TextContent)) {
                continue;
            }
            visibleChildren++;
            if (visibleChildren > 1) return true;
        }
        return false;
    }

    private static bool ContainsMultipleVisibleBlockContentItems(
        IElement element,
        IReadOnlyDictionary<IElement, HtmlComputedStyle> styles) {
        return HasMultipleVisibleBlockContentItems(element, element, styles)
            || element.QuerySelectorAll("*").Any(parent =>
                HasMultipleVisibleBlockContentItems(parent, element, styles));
    }

    private static bool HasMultipleVisibleBlockContentItems(
        IElement parent,
        IElement region,
        IReadOnlyDictionary<IElement, HtmlComputedStyle> styles) {
        int visibleContentItems = 0;
        bool hasBlockItem = false;
        foreach (INode child in parent.ChildNodes) {
            if (child is IElement childElement) {
                if (!IsProjectionElementVisible(childElement, region, styles)
                    || string.IsNullOrWhiteSpace(childElement.TextContent)
                        && childElement.QuerySelector("img") == null) continue;
                if (styles.TryGetValue(childElement, out HtmlComputedStyle? childStyle)
                    && IsBlockBoundaryDisplay(childElement, childStyle.GetValue("display"))) {
                    hasBlockItem = true;
                }
            } else if (string.IsNullOrWhiteSpace(child.TextContent)) {
                continue;
            }
            visibleContentItems++;
            if (hasBlockItem && visibleContentItems > 1) return true;
        }
        return false;
    }

    private static bool IsBlockBoundaryDisplay(IElement element, string value) {
        string display = value.Trim();
        if (display.Length == 0) return HtmlRenderStyleResolver.IsDefaultBlockElement(element);
        return !display.Equals("none", StringComparison.OrdinalIgnoreCase)
            && !display.Equals("contents", StringComparison.OrdinalIgnoreCase)
            && !display.Equals("inline", StringComparison.OrdinalIgnoreCase)
            && !display.Equals("inline-block", StringComparison.OrdinalIgnoreCase)
            && !display.Equals("inline-flex", StringComparison.OrdinalIgnoreCase)
            && !display.Equals("inline-grid", StringComparison.OrdinalIgnoreCase)
            && !display.Equals("inline-table", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsFlexOrGridDisplay(HtmlComputedStyle style) {
        string display = style.GetValue("display").Trim();
        return display.Equals("flex", StringComparison.OrdinalIgnoreCase)
            || display.Equals("inline-flex", StringComparison.OrdinalIgnoreCase)
            || display.Equals("grid", StringComparison.OrdinalIgnoreCase)
            || display.Equals("inline-grid", StringComparison.OrdinalIgnoreCase);
    }

    private static bool ContainsNestedLayoutPlacement(
        IElement element,
        IReadOnlyDictionary<IElement, HtmlComputedStyle> styles,
        bool includeImages) {
        return element.QuerySelectorAll("*").Any(child =>
            (includeImages || child is not IHtmlImageElement)
            && IsProjectionElementVisible(child, element, styles)
            && styles.TryGetValue(child, out HtmlComputedStyle? childStyle)
            && HasLayoutPlacementStyle(childStyle));
    }

    private static bool HasLayoutPlacementStyle(HtmlComputedStyle style) {
        string position = style.GetValue("position").Trim();
        string floatSide = style.GetValue("float").Trim();
        return position.Equals("absolute", StringComparison.OrdinalIgnoreCase)
            || position.Equals("fixed", StringComparison.OrdinalIgnoreCase)
            || floatSide.Equals("left", StringComparison.OrdinalIgnoreCase)
            || floatSide.Equals("right", StringComparison.OrdinalIgnoreCase)
            || IsFlexOrGridDisplay(style);
    }

    private static void AppendInlineContentOrder(
        INode node,
        IElement region,
        IReadOnlyDictionary<IElement, HtmlComputedStyle> styles,
        ICollection<bool> contentOrder) {
        foreach (INode child in node.ChildNodes) {
            if (child is IHtmlImageElement image) {
                if (IsProjectionImageVisible(image, region, styles)) contentOrder.Add(true);
            } else if (child is IElement childElement) {
                AppendInlineContentOrder(childElement, region, styles, contentOrder);
            } else if (!string.IsNullOrWhiteSpace(child.TextContent)) {
                contentOrder.Add(false);
            }
        }
    }

    private static bool TryGetNonNativeRegionEffect(
        IElement element,
        IReadOnlyDictionary<IElement, HtmlComputedStyle> styles,
        out string detail) {
        var effects = new List<string>();
        AddNonNativeBoxModelSpacing(element, styles, effects, "");
        for (IElement? current = element; current != null; current = current.ParentElement) {
            AddNonNativeEffects(current, styles, effects, "");
        }
        foreach (IElement descendant in element.QuerySelectorAll("*")) {
            AddNonNativeEffects(descendant, styles, effects, "descendant:");
            AddNonNativeBoxModelSpacing(descendant, styles, effects, "descendant:");
        }
        detail = string.Join("; ", effects.Distinct(StringComparer.OrdinalIgnoreCase));
        return effects.Count > 0;
    }

    private static void AddNonNativeEffects(
        IElement element,
        IReadOnlyDictionary<IElement, HtmlComputedStyle> styles,
        ICollection<string> effects,
        string prefix) {
        if (!styles.TryGetValue(element, out HtmlComputedStyle? style)) return;
        string opacity = style.GetValue("opacity").Trim();
        if (element is not IHtmlImageElement
            && double.TryParse(opacity, System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out double parsedOpacity)
            && parsedOpacity < 0.999D) {
            effects.Add(prefix + "opacity=" + opacity);
        }
        string backgroundColor = style.GetValue("background-color").Trim();
        if (backgroundColor.Length > 0
            && HtmlRenderCssValues.TryColor(backgroundColor, out OfficeIMO.Drawing.OfficeColor parsedBackground)
            && parsedBackground.A < byte.MaxValue) {
            effects.Add(prefix + "background-color=" + backgroundColor);
        }
        AddNonDefaultEffect(style, "transform", "none", effects, prefix);
        AddNonDefaultEffect(style, "clip-path", "none", effects, prefix);
        AddNonDefaultEffect(style, "filter", "none", effects, prefix);
        AddNonDefaultEffect(style, "mix-blend-mode", "normal", effects, prefix);
        AddOutlineEffect(style, effects, prefix);
        AddNonDefaultEffect(style, "overflow", "visible", effects, prefix);
        AddNonDefaultEffect(style, "overflow-x", "visible", effects, prefix);
        AddNonDefaultEffect(style, "overflow-y", "visible", effects, prefix);
        AddNonZeroEffect(style, "border-radius", effects, prefix);
        AddNonZeroEffect(style, "border-top-left-radius", effects, prefix);
        AddNonZeroEffect(style, "border-top-right-radius", effects, prefix);
        AddNonZeroEffect(style, "border-bottom-right-radius", effects, prefix);
        AddNonZeroEffect(style, "border-bottom-left-radius", effects, prefix);
    }

    private static void AddOutlineEffect(
        HtmlComputedStyle style,
        ICollection<string> effects,
        string prefix) {
        string shorthand = style.GetValue("outline").Trim();
        string widthValue = style.GetValue("outline-width").Trim();
        string styleValue = style.GetValue("outline-style").Trim();
        string colorValue = style.GetValue("outline-color").Trim();
        string offsetValue = style.GetValue("outline-offset").Trim();
        if (shorthand.Length == 0 && widthValue.Length == 0 && styleValue.Length == 0
            && colorValue.Length == 0 && offsetValue.Length == 0) return;
        if (!HtmlCssBoxStrokeParser.TryParseOutline(
                style, 100D, 16D, 16D, 100D, 100D,
                OfficeIMO.Drawing.OfficeColor.Black,
                out double width, out string outlineStyle, out _, out _, out string detail)) {
            effects.Add(prefix + (detail.Length > 0 ? detail : "outline=unsupported"));
            return;
        }
        if (width <= 0D || outlineStyle.Equals("none", StringComparison.OrdinalIgnoreCase)
            || outlineStyle.Equals("hidden", StringComparison.OrdinalIgnoreCase)) return;
        effects.Add(prefix + (shorthand.Length > 0
            ? "outline=" + shorthand
            : "outline-style=" + outlineStyle + "; outline-width=" + widthValue));
    }

    private static void AddNonNativeBoxModelSpacing(
        IElement element,
        IReadOnlyDictionary<IElement, HtmlComputedStyle> styles,
        ICollection<string> effects,
        string prefix) {
        if (!styles.TryGetValue(element, out HtmlComputedStyle? style)) return;
        AddNonZeroEffect(style, "padding", effects, prefix);
        AddNonZeroEffect(style, "padding-top", effects, prefix);
        AddNonZeroEffect(style, "padding-right", effects, prefix);
        AddNonZeroEffect(style, "padding-bottom", effects, prefix);
        AddNonZeroEffect(style, "padding-left", effects, prefix);
        AddNonZeroEffect(style, "padding-block", effects, prefix);
        AddNonZeroEffect(style, "padding-block-start", effects, prefix);
        AddNonZeroEffect(style, "padding-block-end", effects, prefix);
        AddNonZeroEffect(style, "padding-inline", effects, prefix);
        AddNonZeroEffect(style, "padding-inline-start", effects, prefix);
        AddNonZeroEffect(style, "padding-inline-end", effects, prefix);
        AddNonZeroEffect(style, "margin", effects, prefix);
        AddNonZeroEffect(style, "margin-top", effects, prefix);
        AddNonZeroEffect(style, "margin-right", effects, prefix);
        AddNonZeroEffect(style, "margin-bottom", effects, prefix);
        AddNonZeroEffect(style, "margin-left", effects, prefix);
        AddNonZeroEffect(style, "margin-block", effects, prefix);
        AddNonZeroEffect(style, "margin-block-start", effects, prefix);
        AddNonZeroEffect(style, "margin-block-end", effects, prefix);
        AddNonZeroEffect(style, "margin-inline", effects, prefix);
        AddNonZeroEffect(style, "margin-inline-start", effects, prefix);
        AddNonZeroEffect(style, "margin-inline-end", effects, prefix);
    }

    private static void AddNonZeroEffect(
        HtmlComputedStyle style,
        string property,
        ICollection<string> effects,
        string prefix) {
        string value = style.GetValue(property).Trim();
        if (value.Length > 0 && !IsZeroCssValue(value)) effects.Add(prefix + property + "=" + value);
    }

    private static bool IsZeroCssValue(string value) {
        string[] tokens = value.Replace("/", " ").Split(
            new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
        if (tokens.Length == 0) return true;
        foreach (string token in tokens) {
            int length = 0;
            while (length < token.Length && (char.IsDigit(token[length])
                    || token[length] == '+' || token[length] == '-' || token[length] == '.')) {
                length++;
            }
            if (length == 0
                || !double.TryParse(token.Substring(0, length),
                    System.Globalization.NumberStyles.Float,
                    System.Globalization.CultureInfo.InvariantCulture,
                    out double number)
                || Math.Abs(number) > 0.000001D) {
                return false;
            }
        }
        return true;
    }

    private static void AddNonDefaultEffect(
        HtmlComputedStyle style,
        string property,
        string defaultValue,
        ICollection<string> effects,
        string prefix = "") {
        string value = style.GetValue(property).Trim();
        if (value.Length > 0 && !value.Equals(defaultValue, StringComparison.OrdinalIgnoreCase)) {
            effects.Add(prefix + property + "=" + value);
        }
    }

    private static bool HasDistinctRichTextStyle(
        IElement element,
        IReadOnlyDictionary<IElement, HtmlComputedStyle> styles) =>
        HasDistinctStyle(element, styles, RichTextStyleProperties);

    private static bool HasInheritedRichTextStyle(
        IElement element,
        IReadOnlyDictionary<IElement, HtmlComputedStyle> styles) =>
        styles.TryGetValue(element, out HtmlComputedStyle? style)
        && RichTextStyleProperties.Any(property => style.IsInheritedValue(property)
            && !string.IsNullOrWhiteSpace(style.GetValue(property)));

    private static bool HasVisibleRootBorder(
        IElement element,
        IReadOnlyDictionary<IElement, HtmlComputedStyle> styles) =>
        styles.TryGetValue(element, out HtmlComputedStyle? style)
        && HtmlCssBoxStrokeParser.HasBorderDeclaration(style);

    private static bool HasDistinctStyle(
        IElement element,
        IReadOnlyDictionary<IElement, HtmlComputedStyle> styles,
        IReadOnlyList<string> properties) {
        if (!styles.TryGetValue(element, out HtmlComputedStyle? style)
            || element.ParentElement == null
            || !styles.TryGetValue(element.ParentElement, out HtmlComputedStyle? parentStyle)) return false;
        return properties.Any(property => !style.IsInheritedValue(property)
            && !style.IsResetValue(property)
            && !string.IsNullOrWhiteSpace(style.GetValue(property))
            && !string.Equals(style.GetValue(property), parentStyle.GetValue(property),
                StringComparison.OrdinalIgnoreCase));
    }
}
