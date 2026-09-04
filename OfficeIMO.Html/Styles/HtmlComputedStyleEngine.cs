using AngleSharp.Css.Parser;
using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using System.Globalization;

namespace OfficeIMO.Html;

/// <summary>
/// Lightweight computed-style helper for OfficeIMO conversion diagnostics and contract tests.
/// </summary>
public static partial class HtmlComputedStyleEngine {
    private static readonly HashSet<string> InheritedProperties = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
        "caption-side",
        "border-collapse",
        "border-spacing",
        "color",
        "direction",
        "font-family",
        "font-size",
        "font-style",
        "font-variant",
        "font-variant-caps",
        "font-weight",
        "hyphens",
        "hyphenate-character",
        "hyphenate-limit-chars",
        "hyphenate-limit-last",
        "hyphenate-limit-lines",
        "hyphenate-limit-zone",
        "image-orientation",
        "image-resolution",
        "line-height",
        "letter-spacing",
        "list-style",
        "list-style-image",
        "list-style-position",
        "list-style-type",
        "tab-size",
        "orphans",
        "page",
        "quotes",
        "text-align",
        "text-transform",
        "visibility",
        "widows",
        "white-space",
        "word-spacing"
    };
    private static readonly HashSet<string> SupportedProperties = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
        "background",
        "background-attachment",
        "background-clip",
        "background-color",
        "background-image",
        "background-origin",
        "background-position",
        "background-repeat",
        "background-size",
        "align-content",
        "align-items",
        "align-self",
        "animation",
        "animation-name",
        "aspect-ratio",
        "bottom",
        "bookmark-label",
        "bookmark-level",
        "bookmark-state",
        "box-decoration-break",
        "box-shadow",
        "border",
        "border-bottom",
        "border-bottom-color",
        "border-bottom-style",
        "border-bottom-width",
        "border-collapse",
        "border-spacing",
        "border-color",
        "border-left",
        "border-left-color",
        "border-left-style",
        "border-left-width",
        "border-right",
        "border-right-color",
        "border-right-style",
        "border-right-width",
        "border-radius",
        "border-style",
        "border-top",
        "border-top-color",
        "border-top-left-radius",
        "border-top-right-radius",
        "border-top-style",
        "border-top-width",
        "border-bottom-left-radius",
        "border-bottom-right-radius",
        "border-width",
        "box-sizing",
        "break-after",
        "break-before",
        "break-inside",
        "caption-side",
        "clear",
        "clip-path",
        "color",
        "column-gap",
        "column-count",
        "column-fill",
        "column-rule",
        "column-rule-color",
        "column-rule-style",
        "column-rule-width",
        "column-span",
        "column-width",
        "columns",
        "content",
        "container",
        "container-name",
        "container-type",
        "counter-increment",
        "counter-reset",
        "counter-set",
        "cursor",
        "direction",
        "display",
        "font-family",
        "font-size",
        "font-style",
        "font-variant",
        "font-variant-caps",
        "font-weight",
        "flex",
        "flex-basis",
        "flex-direction",
        "flex-flow",
        "flex-grow",
        "flex-shrink",
        "flex-wrap",
        "filter",
        "float",
        "gap",
        "grid-area",
        "grid-auto-columns",
        "grid-auto-flow",
        "grid-auto-rows",
        "grid-column",
        "grid-column-end",
        "grid-column-start",
        "grid-row",
        "grid-row-end",
        "grid-row-start",
        "grid-template-columns",
        "grid-template-areas",
        "grid-template-rows",
        "height",
        "hyphens",
        "hyphenate-character",
        "hyphenate-limit-chars",
        "hyphenate-limit-last",
        "hyphenate-limit-lines",
        "hyphenate-limit-zone",
        "image-orientation",
        "image-resolution",
        "left",
        "letter-spacing",
        "line-height",
        "list-style",
        "list-style-image",
        "list-style-position",
        "list-style-type",
        "justify-content",
        "justify-items",
        "justify-self",
        "margin",
        "margin-bottom",
        "margin-left",
        "margin-right",
        "margin-top",
        "max-height",
        "max-width",
        "min-height",
        "min-width",
        "object-fit",
        "object-position",
        "opacity",
        "order",
        "orphans",
        "outline-color",
        "outline",
        "outline-offset",
        "outline-style",
        "outline-width",
        "overflow",
        "overflow-clip-margin",
        "overflow-x",
        "overflow-y",
        "overflow-wrap",
        "page",
        "page-break-after",
        "page-break-before",
        "page-break-inside",
        "padding",
        "padding-bottom",
        "padding-left",
        "padding-right",
        "padding-top",
        "position",
        "quotes",
        "place-content",
        "place-items",
        "place-self",
        "right",
        "row-gap",
        "string-set",
        "tab-size",
        "text-align",
        "text-indent",
        "text-decoration-line",
        "text-decoration-color",
        "text-decoration-style",
        "text-overflow",
        "text-transform",
        "table-layout",
        "transform",
        "transform-origin",
        "top",
        "vertical-align",
        "visibility",
        "white-space",
        "width",
        "widows",
        "word-break",
        "word-spacing",
        "line-clamp",
        "-webkit-line-clamp",
        "-officeimo-pdf-tag-type",
        "z-index"
    };

    /// <summary>
    /// Computes styles for every element in the supplied document using style tags and inline style attributes.
    /// </summary>
    public static IReadOnlyDictionary<IElement, HtmlComputedStyle> Compute(IHtmlDocument document, HtmlCssMediaContext mediaContext = HtmlCssMediaContext.Screen) {
        return ComputeStyleSet(document, MediaEnvironment.CreateDefault(mediaContext), false, limits: null).Elements;
    }

    internal static IReadOnlyDictionary<IElement, HtmlComputedStyle> Compute(
        IHtmlDocument document,
        HtmlCssMediaContext mediaContext,
        HtmlConversionLimits limits) =>
        ComputeStyleSet(document, MediaEnvironment.CreateDefault(mediaContext), false, limits).Elements;

    internal static IReadOnlyDictionary<IElement, HtmlComputedStyle> Compute(
        IHtmlDocument document,
        HtmlResourcePipelineOptions options) {
        MediaEnvironment environment = options.MediaWidth.HasValue && options.MediaHeight.HasValue
            ? new MediaEnvironment(options.MediaContext, options.MediaWidth.Value, options.MediaHeight.Value, options.MediaFeatures)
            : MediaEnvironment.CreateDefault(options.MediaContext, options.MediaFeatures);
        return ComputeStyleSet(document, environment, false, options.Limits).Elements;
    }

    internal static HtmlComputedStyleSet ComputeForProvenance(
        IHtmlDocument document,
        HtmlResourcePipelineOptions options) {
        MediaEnvironment environment = options.MediaWidth.HasValue && options.MediaHeight.HasValue
            ? new MediaEnvironment(options.MediaContext, options.MediaWidth.Value, options.MediaHeight.Value, options.MediaFeatures)
            : MediaEnvironment.CreateDefault(options.MediaContext, options.MediaFeatures);
        return ComputeStyleSet(document, environment, true, options.Limits);
    }

    internal static HtmlComputedStyleSet ComputeForRendering(IHtmlDocument document, HtmlRenderOptions options, HtmlConversionLimits limits) =>
        ComputeStyleSet(
            document,
            new MediaEnvironment(
                options.MediaContext,
                options.Mode == HtmlRenderMode.Paged ? options.PageWidth : options.ViewportWidth,
                options.Mode == HtmlRenderMode.Paged ? options.PageHeight : options.ViewportHeight ?? 1056D,
                options.MediaFeatures),
            true,
            limits);

    private static HtmlComputedStyleSet ComputeStyleSet(
        IHtmlDocument document,
        MediaEnvironment environment,
        bool includePseudoElements,
        HtmlConversionLimits? limits) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        var budget = new HtmlCssProcessingBudget(limits);
        IReadOnlyList<StyleRule> rules = ParseStyleRules(document, environment, budget);
        var ruleIndex = new StyleRuleIndex(rules);
        var computed = new Dictionary<IElement, HtmlComputedStyle>();
        var pseudoElements = new Dictionary<IElement, HtmlPseudoElementStylePair>();
        IElement? root = document.DocumentElement ?? document.Body;
        if (root != null) {
            ComputeElement(root, null, ruleIndex, computed, pseudoElements, includePseudoElements, budget, environment, environment.Width, environment.Height, Array.Empty<ContainerQueryContext>());
        }

        return new HtmlComputedStyleSet(computed, pseudoElements);
    }

    /// <summary>
    /// Parses raw HTML through the bounded shared conversion document and computes styles for matching elements.
    /// </summary>
    public static IReadOnlyDictionary<IElement, HtmlComputedStyle> Compute(string html, HtmlCssMediaContext mediaContext = HtmlCssMediaContext.Screen) {
        return Compute(HtmlConversionDocument.Parse(html), mediaContext);
    }

    /// <summary>Computes styles from a retained conversion document without reparsing its HTML source.</summary>
    public static IReadOnlyDictionary<IElement, HtmlComputedStyle> Compute(
        HtmlConversionDocument document,
        HtmlCssMediaContext mediaContext = HtmlCssMediaContext.Screen) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        return Compute(document.CreateSourceDocumentForConversion(), mediaContext, document.Limits);
    }

    /// <summary>
    /// Creates a compact summary from computed style results.
    /// </summary>
    public static HtmlComputedStyleSummary Summarize(IReadOnlyDictionary<IElement, HtmlComputedStyle> styles) {
        if (styles == null) {
            throw new ArgumentNullException(nameof(styles));
        }

        var propertyNames = new List<string>();
        var fontFamilies = new List<string>();
        var colorValues = new List<string>();
        int styledElementCount = 0;
        int hiddenElementCount = 0;
        foreach (HtmlComputedStyle style in styles.Values) {
            if (style.Properties.Count > 0) {
                styledElementCount++;
            }

            if (IsEffectivelyHidden(style)) {
                hiddenElementCount++;
            }

            foreach (KeyValuePair<string, string> pair in style.Properties) {
                propertyNames.Add(pair.Key);
                if (string.Equals(pair.Key, "font-family", StringComparison.OrdinalIgnoreCase)) {
                    fontFamilies.Add(pair.Value);
                }

                if (IsColorProperty(pair.Key)) {
                    colorValues.Add(pair.Value);
                }
            }
        }

        return new HtmlComputedStyleSummary(
            styles.Count,
            styledElementCount,
            hiddenElementCount,
            propertyNames,
            fontFamilies,
            colorValues);
    }

    private static void ComputeElement(
        IElement element,
        HtmlComputedStyle? parent,
        StyleRuleIndex rules,
        IDictionary<IElement, HtmlComputedStyle> computed,
        IDictionary<IElement, HtmlPseudoElementStylePair> pseudoElements,
        bool includePseudoElements,
        HtmlCssProcessingBudget budget,
        MediaEnvironment environment,
        double containingWidth,
        double? containingHeight,
        IReadOnlyList<ContainerQueryContext> containerContexts) {
        var properties = new Dictionary<string, CascadedProperty>(HtmlCssPropertyNameComparer.Instance);

        string? directionAttribute = element.GetAttribute("dir")?.Trim();
        if (string.Equals(directionAttribute, "ltr", StringComparison.OrdinalIgnoreCase)
            || string.Equals(directionAttribute, "rtl", StringComparison.OrdinalIgnoreCase)) {
            properties["direction"] = new CascadedProperty(directionAttribute!.ToLowerInvariant(), false, Specificity.PresentationalHint, -1);
        }

        IReadOnlyList<StyleRule> candidateRules = rules.GetCandidates(element);
        foreach (StyleRule rule in candidateRules) {
            budget.RecordSelectorEvaluation();
            if (AreContainerConditionsApplicable(rule.ContainerConditions, containerContexts, environment)
                && !TryParsePseudoElementSelector(rule.Selector, out _, out _)
                && MatchesSelector(element, rule.Selector)) {
                foreach (var declaration in rule.Declarations) {
                    if (declaration.Value.IsSupported) {
                        ApplyDeclaration(properties, parent?.Properties, declaration.Key, declaration.Value.Value, declaration.Value.IsImportant, rule.Specificity, rule.Order, rule.LayerOrder, valueAlreadyValidated: true);
                    }
                }
            }
        }

        ApplyInlineDeclarations(properties, parent?.Properties, element.GetAttribute("style"));
        Dictionary<string, string> resolvedProperties = ResolveComputedProperties(properties, parent?.Properties,
            out HashSet<string> inheritedProperties, out HashSet<string> resetProperties,
            out HashSet<string> specifiedProperties);
        HtmlComputedStyle style = HtmlComputedStyle.FromOwnedCollections(
            resolvedProperties, inheritedProperties, resetProperties, specifiedProperties);
        computed[element] = style;

        double inheritedFontSize = containerContexts.Count == 0 ? 16D : containerContexts[containerContexts.Count - 1].FontSize;
        double rootFontSize = containerContexts.Count == 0 ? 16D : containerContexts[0].RootFontSize;
        ResolveContainerUnitDimensions(containerContexts, out double containerUnitWidth, out double containerUnitHeight);
        double elementFontSize = ResolveContainerFontSize(style, environment, inheritedFontSize, rootFontSize, containerUnitWidth, containerUnitHeight);
        style.ResolvedFontSizePoints = elementFontSize * 72D / 96D;
        if (containerContexts.Count == 0) rootFontSize = elementFontSize;
        double elementWidth = ResolveContainerElementWidth(style, containingWidth, elementFontSize, rootFontSize, environment, containerUnitWidth, containerUnitHeight);
        double? elementHeight = ResolveContainerElementHeight(style, elementWidth, containingWidth, containingHeight, elementFontSize, rootFontSize, environment, containerUnitWidth, containerUnitHeight);
        IReadOnlyList<ContainerQueryContext> childContainerContexts = AddContainerContext(style, elementWidth, elementHeight, elementFontSize, inheritedFontSize, rootFontSize, containerContexts);
        if (includePseudoElements) ComputePseudoElementStyles(element, style, candidateRules, pseudoElements, budget, childContainerContexts, environment);

        foreach (IElement child in element.Children) {
            ComputeElement(child, style, rules, computed, pseudoElements, includePseudoElements, budget, environment, elementWidth, elementHeight, childContainerContexts);
        }
    }

    private static void ComputePseudoElementStyles(
        IElement element,
        HtmlComputedStyle originatingStyle,
        IReadOnlyList<StyleRule> candidateRules,
        IDictionary<IElement, HtmlPseudoElementStylePair> pseudoElements,
        HtmlCssProcessingBudget budget,
        IReadOnlyList<ContainerQueryContext> containerContexts,
        MediaEnvironment environment) {
        HtmlComputedStyle? before = ComputePseudoElementStyle(element, originatingStyle, candidateRules, HtmlPseudoElementKind.Before, budget, containerContexts, environment);
        HtmlComputedStyle? after = ComputePseudoElementStyle(element, originatingStyle, candidateRules, HtmlPseudoElementKind.After, budget, containerContexts, environment);
        HtmlComputedStyle? marker = ComputePseudoElementStyle(element, originatingStyle, candidateRules, HtmlPseudoElementKind.Marker, budget, containerContexts, environment);
        if (before == null && after == null && marker == null) return;
        pseudoElements[element] = new HtmlPseudoElementStylePair { Before = before, After = after, Marker = marker };
    }

    private static HtmlComputedStyle? ComputePseudoElementStyle(
        IElement element,
        HtmlComputedStyle originatingStyle,
        IReadOnlyList<StyleRule> candidateRules,
        HtmlPseudoElementKind kind,
        HtmlCssProcessingBudget budget,
        IReadOnlyList<ContainerQueryContext> containerContexts,
        MediaEnvironment environment) {
        List<StyleRule>? matchedRules = null;
        foreach (StyleRule rule in candidateRules) {
            budget.RecordSelectorEvaluation();
            if (!AreContainerConditionsApplicable(rule.ContainerConditions, containerContexts, environment)
                || !TryParsePseudoElementSelector(rule.Selector, out string hostSelector, out HtmlPseudoElementKind ruleKind)
                || ruleKind != kind
                || !MatchesSelector(element, hostSelector)) {
                continue;
            }

            (matchedRules ??= new List<StyleRule>()).Add(rule);
        }

        if (matchedRules == null) return null;
        var properties = new Dictionary<string, CascadedProperty>(HtmlCssPropertyNameComparer.Instance);

        foreach (StyleRule rule in matchedRules) {
            foreach (KeyValuePair<string, StyleDeclaration> declaration in rule.Declarations) {
                if (!declaration.Value.IsSupported) continue;
                ApplyDeclaration(
                    properties,
                    originatingStyle.Properties,
                    declaration.Key,
                    declaration.Value.Value,
                    declaration.Value.IsImportant,
                    rule.Specificity,
                    rule.Order,
                    rule.LayerOrder,
                    valueAlreadyValidated: true);
            }
        }

        Dictionary<string, string> resolvedProperties = ResolveComputedProperties(properties, originatingStyle.Properties,
            out HashSet<string> inheritedProperties, out HashSet<string> resetProperties,
            out HashSet<string> specifiedProperties);
        return HtmlComputedStyle.FromOwnedCollections(
            resolvedProperties, inheritedProperties, resetProperties, specifiedProperties);
    }

}
