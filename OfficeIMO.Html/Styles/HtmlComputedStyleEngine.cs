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
        "unicode-bidi",
        "font-family",
        "font-size",
        "font-stretch",
        "font-style",
        "font-variant",
        "font-variant-caps",
        "font-variant-east-asian",
        "font-variant-ligatures",
        "font-variant-numeric",
        "font-feature-settings",
        "font-kerning",
        "font-palette",
        "font-weight",
        "fill",
        "fill-opacity",
        "fill-rule",
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
        "pointer-events",
        "quotes",
        "text-align",
        "text-shadow",
        "text-transform",
        "visibility",
        "widows",
        "white-space",
        "word-spacing",
        "writing-mode",
        "text-orientation",
        "ruby-position",
        "ruby-align",
        "stroke",
        "stroke-width",
        "stroke-opacity",
        "stroke-dasharray",
        "stroke-dashoffset",
        "stroke-linecap",
        "stroke-linejoin",
        "stroke-miterlimit",
        "marker-start",
        "marker-mid",
        "marker-end",
        "text-anchor",
        "dominant-baseline",
        "baseline-shift"
    };
    private static readonly HashSet<string> SupportedProperties = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
        "all",
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
        "font",
        "font-family",
        "font-size",
        "font-stretch",
        "font-style",
        "font-variant",
        "font-variant-caps",
        "font-variant-east-asian",
        "font-variant-ligatures",
        "font-variant-numeric",
        "font-feature-settings",
        "font-kerning",
        "font-palette",
        "font-weight",
        "fill",
        "fill-opacity",
        "fill-rule",
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
        "marker-start",
        "marker-mid",
        "marker-end",
        "mask",
        "mix-blend-mode",
        "list-style",
        "list-style-image",
        "list-style-position",
        "list-style-type",
        "justify-content",
        "justify-items",
        "justify-self",
        "margin",
        "margin-bottom",
        "margin-block",
        "margin-block-end",
        "margin-block-start",
        "margin-inline",
        "margin-inline-end",
        "margin-inline-start",
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
        "pointer-events",
        "padding",
        "padding-bottom",
        "padding-block",
        "padding-block-end",
        "padding-block-start",
        "padding-inline",
        "padding-inline-end",
        "padding-inline-start",
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
        "text-shadow",
        "text-transform",
        "table-layout",
        "stroke",
        "stroke-width",
        "stroke-opacity",
        "stroke-dasharray",
        "stroke-dashoffset",
        "stroke-linecap",
        "stroke-linejoin",
        "stroke-miterlimit",
        "text-anchor",
        "dominant-baseline",
        "baseline-shift",
        "transform",
        "transform-origin",
        "top",
        "inset-block",
        "inset-block-end",
        "inset-block-start",
        "inset-inline",
        "inset-inline-end",
        "inset-inline-start",
        "block-size",
        "inline-size",
        "min-block-size",
        "min-inline-size",
        "max-block-size",
        "max-inline-size",
        "border-block",
        "border-block-color",
        "border-block-end",
        "border-block-end-color",
        "border-block-end-style",
        "border-block-end-width",
        "border-block-start",
        "border-block-start-color",
        "border-block-start-style",
        "border-block-start-width",
        "border-block-style",
        "border-block-width",
        "border-inline",
        "border-inline-color",
        "border-inline-end",
        "border-inline-end-color",
        "border-inline-end-style",
        "border-inline-end-width",
        "border-inline-start",
        "border-inline-start-color",
        "border-inline-start-style",
        "border-inline-start-width",
        "border-inline-style",
        "border-inline-width",
        "border-end-end-radius",
        "border-end-start-radius",
        "border-start-end-radius",
        "border-start-start-radius",
        "vertical-align",
        "visibility",
        "white-space",
        "width",
        "widows",
        "word-break",
        "word-spacing",
        "writing-mode",
        "text-orientation",
        "ruby-position",
        "ruby-align",
        "line-clamp",
        "-webkit-line-clamp",
        "-officeimo-pdf-tag-type",
        "z-index"
    };

    /// <summary>
    /// Computes styles for every element in the supplied document using style tags and inline style attributes.
    /// </summary>
    internal static IReadOnlyDictionary<IElement, HtmlComputedStyle> Compute(IHtmlDocument document, HtmlCssMediaContext mediaContext = HtmlCssMediaContext.Screen) {
        return ComputeStyleSet(document, MediaEnvironment.CreateDefault(mediaContext), false, false, limits: null).Elements;
    }

    internal static IReadOnlyDictionary<IElement, HtmlComputedStyle> Compute(
        IHtmlDocument document,
        HtmlCssMediaContext mediaContext,
        HtmlConversionLimits limits) =>
        ComputeStyleSet(document, MediaEnvironment.CreateDefault(mediaContext), false, false, limits).Elements;

    internal static IReadOnlyDictionary<IElement, HtmlComputedStyle> Compute(
        IHtmlDocument document,
        HtmlResourcePipelineOptions options) {
        MediaEnvironment environment = options.MediaWidth.HasValue && options.MediaHeight.HasValue
            ? new MediaEnvironment(options.MediaContext, options.MediaWidth.Value, options.MediaHeight.Value, options.MediaFeatures)
            : MediaEnvironment.CreateDefault(options.MediaContext, options.MediaFeatures);
        return ComputeStyleSet(document, environment, false, false, options.Limits).Elements;
    }

    internal static HtmlComputedStyleSet ComputeForProvenance(
        IHtmlDocument document,
        HtmlResourcePipelineOptions options) {
        MediaEnvironment environment = options.MediaWidth.HasValue && options.MediaHeight.HasValue
            ? new MediaEnvironment(options.MediaContext, options.MediaWidth.Value, options.MediaHeight.Value, options.MediaFeatures)
            : MediaEnvironment.CreateDefault(options.MediaContext, options.MediaFeatures);
        return ComputeStyleSet(document, environment, true, false, options.Limits);
    }

    internal static HtmlComputedStyleSet ComputeForContentSafety(
        IHtmlDocument document,
        HtmlConversionLimits? limits = null) =>
        ComputeStyleSet(document, MediaEnvironment.CreateDefault(HtmlCssMediaContext.Screen), true, false, limits);

    internal static HtmlComputedStyleSet ComputeForRendering(IHtmlDocument document, HtmlRenderOptions options, HtmlConversionLimits limits) =>
        ComputeStyleSet(
            document,
            new MediaEnvironment(
                options.MediaContext,
                options.Mode == HtmlRenderMode.Paged ? options.PageWidth : options.ViewportWidth,
                options.Mode == HtmlRenderMode.Paged ? options.PageHeight : options.ViewportHeight ?? 1056D,
                options.MediaFeatures),
            true,
            false,
            limits);

    private static HtmlComputedStyleSet ComputeStyleSet(
        IHtmlDocument document,
        MediaEnvironment environment,
        bool includePseudoElements,
        bool includeCascadeTraces,
        HtmlConversionLimits? limits) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        var budget = new HtmlCssProcessingBudget(limits, MatchProviderSelectorFragment);
        IReadOnlyDictionary<string, CustomPropertyRegistration> customPropertyRegistrations =
            ParseCustomPropertyRegistrations(document, environment);
        IReadOnlyList<StyleRule> rules = ParseStyleRules(document, environment, budget);
        var ruleIndex = new StyleRuleIndex(rules, customPropertyRegistrations);
        var computed = new Dictionary<IElement, HtmlComputedStyle>();
        var pseudoElements = new Dictionary<IElement, HtmlPseudoElementStylePair>();
        IElement? root = document.DocumentElement ?? document.Body;
        if (root != null) {
            ComputeElement(root, null, ruleIndex, computed, pseudoElements, includePseudoElements, includeCascadeTraces,
                budget, environment, environment.Width, environment.Height, Array.Empty<ContainerQueryContext>());
        }

        return new HtmlComputedStyleSet(computed, pseudoElements);
    }

    /// <summary>
    /// Creates a compact summary from computed style results.
    /// </summary>
    public static HtmlComputedStyleSummary Summarize(IReadOnlyDictionary<Dom.HtmlElement, HtmlComputedStyle> styles) =>
        SummarizeValues((styles ?? throw new ArgumentNullException(nameof(styles))).Values, styles.Count);

    internal static HtmlComputedStyleSummary Summarize(IReadOnlyDictionary<IElement, HtmlComputedStyle> styles) =>
        SummarizeValues((styles ?? throw new ArgumentNullException(nameof(styles))).Values, styles.Count);

    private static HtmlComputedStyleSummary SummarizeValues(IEnumerable<HtmlComputedStyle> styles, int count) {

        var propertyNames = new List<string>();
        var fontFamilies = new List<string>();
        var colorValues = new List<string>();
        int styledElementCount = 0;
        int hiddenElementCount = 0;
        foreach (HtmlComputedStyle style in styles) {
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
            count,
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
        bool includeCascadeTraces,
        HtmlCssProcessingBudget budget,
        MediaEnvironment environment,
        double containingWidth,
        double? containingHeight,
        IReadOnlyList<ContainerQueryContext> containerContexts) {
        var properties = new Dictionary<string, CascadedProperty>(HtmlCssPropertyNameComparer.Instance);

        string? directionAttribute = element.GetAttribute("dir")?.Trim();
        if (string.Equals(directionAttribute, "ltr", StringComparison.OrdinalIgnoreCase)
            || string.Equals(directionAttribute, "rtl", StringComparison.OrdinalIgnoreCase)) {
            properties["direction"] = new CascadedProperty(directionAttribute!.ToLowerInvariant(), false,
                Specificity.PresentationalHint, -1, source: OfficeIMO.Html.Css.HtmlCssCascadeSourceKind.PresentationalHint);
        } else if (string.Equals(directionAttribute, "auto", StringComparison.OrdinalIgnoreCase)
            || string.Equals(element.TagName, "bdi", StringComparison.OrdinalIgnoreCase) && string.IsNullOrEmpty(directionAttribute)) {
            OfficeIMO.Drawing.OfficeTextDirection resolvedDirection =
                OfficeIMO.Drawing.OfficeTextElements.ResolveBaseDirection(element.TextContent);
            properties["direction"] = new CascadedProperty(
                resolvedDirection == OfficeIMO.Drawing.OfficeTextDirection.RightToLeft ? "rtl" : "ltr",
                false,
                Specificity.PresentationalHint,
                -1,
                source: OfficeIMO.Html.Css.HtmlCssCascadeSourceKind.PresentationalHint);
        }
        if (string.Equals(element.TagName, "bdi", StringComparison.OrdinalIgnoreCase)) {
            properties["unicode-bidi"] = new CascadedProperty("isolate", false, Specificity.PresentationalHint, -1,
                source: OfficeIMO.Html.Css.HtmlCssCascadeSourceKind.PresentationalHint);
        } else if (string.Equals(element.TagName, "bdo", StringComparison.OrdinalIgnoreCase)) {
            properties["unicode-bidi"] = new CascadedProperty("bidi-override", false, Specificity.PresentationalHint, -1,
                source: OfficeIMO.Html.Css.HtmlCssCascadeSourceKind.PresentationalHint);
        }

        IReadOnlyList<StyleRule> candidateRules = rules.GetCandidates(element);
        foreach (StyleRule rule in candidateRules) {
            budget.RecordSelectorEvaluation();
            if (AreContainerConditionsApplicable(rule.ContainerConditions, containerContexts, environment)
                && !TryParsePseudoElementSelector(rule.Selector, out _, out _)
                && MatchesSelector(element, rule, budget)) {
                foreach (var declaration in rule.Declarations) {
                    if (declaration.Value.IsSupported) {
                        ApplyDeclaration(properties, parent?.Properties, declaration.Key, declaration.Value.Value,
                            declaration.Value.IsImportant, rule.Specificity, rule.Order, rule.LayerOrder,
                            valueAlreadyValidated: true, declarationOrder: declaration.Value.DeclarationOrder,
                            customPropertyRegistrations: rules.CustomPropertyRegistrations,
                            source: OfficeIMO.Html.Css.HtmlCssCascadeSourceKind.StyleRule,
                            selector: rule.Selector, layerName: rule.LayerName);
                    }
                }
            }
        }

        ApplyInlineDeclarations(properties, parent?.Properties, element.GetAttribute("style"), rules.CustomPropertyRegistrations, budget);
        Dictionary<string, string> resolvedProperties = ResolveComputedProperties(properties, parent?.Properties,
            out HashSet<string> inheritedProperties, out HashSet<string> resetProperties,
            out HashSet<string> originRevertedProperties,
            out HashSet<string> specifiedProperties,
            out Dictionary<string, HtmlCssCascadePriority> cascadePriorities,
            out Dictionary<string, OfficeIMO.Html.Css.HtmlCssCascadeTrace>? cascadeTraces,
            includeCascadeTraces,
            rules.CustomPropertyRegistrations);
        HtmlComputedStyle style = HtmlComputedStyle.FromOwnedCollections(
            resolvedProperties, inheritedProperties, resetProperties, originRevertedProperties,
            specifiedProperties, cascadePriorities, cascadeTraces);
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
        if (includePseudoElements) ComputePseudoElementStyles(element, style, candidateRules, pseudoElements, budget,
            childContainerContexts, environment, rules.CustomPropertyRegistrations, includeCascadeTraces);

        foreach (IElement child in element.Children) {
            ComputeElement(child, style, rules, computed, pseudoElements, includePseudoElements, includeCascadeTraces,
                budget, environment, elementWidth, elementHeight, childContainerContexts);
        }
    }

    private static void ComputePseudoElementStyles(
        IElement element,
        HtmlComputedStyle originatingStyle,
        IReadOnlyList<StyleRule> candidateRules,
        IDictionary<IElement, HtmlPseudoElementStylePair> pseudoElements,
        HtmlCssProcessingBudget budget,
        IReadOnlyList<ContainerQueryContext> containerContexts,
        MediaEnvironment environment,
        IReadOnlyDictionary<string, CustomPropertyRegistration> customPropertyRegistrations,
        bool includeCascadeTraces) {
        HtmlComputedStyle? before = ComputePseudoElementStyle(element, originatingStyle, candidateRules, HtmlPseudoElementKind.Before, budget, containerContexts, environment, customPropertyRegistrations, includeCascadeTraces);
        HtmlComputedStyle? after = ComputePseudoElementStyle(element, originatingStyle, candidateRules, HtmlPseudoElementKind.After, budget, containerContexts, environment, customPropertyRegistrations, includeCascadeTraces);
        HtmlComputedStyle? marker = ComputePseudoElementStyle(element, originatingStyle, candidateRules, HtmlPseudoElementKind.Marker, budget, containerContexts, environment, customPropertyRegistrations, includeCascadeTraces);
        HtmlComputedStyle? footnoteCall = ComputePseudoElementStyle(element, originatingStyle, candidateRules, HtmlPseudoElementKind.FootnoteCall, budget, containerContexts, environment, customPropertyRegistrations, includeCascadeTraces);
        HtmlComputedStyle? footnoteMarker = ComputePseudoElementStyle(element, originatingStyle, candidateRules, HtmlPseudoElementKind.FootnoteMarker, budget, containerContexts, environment, customPropertyRegistrations, includeCascadeTraces);
        HtmlComputedStyle? firstLetter = ComputePseudoElementStyle(element, originatingStyle, candidateRules, HtmlPseudoElementKind.FirstLetter, budget, containerContexts, environment, customPropertyRegistrations, includeCascadeTraces);
        HtmlComputedStyle? firstLine = ComputePseudoElementStyle(element, originatingStyle, candidateRules, HtmlPseudoElementKind.FirstLine, budget, containerContexts, environment, customPropertyRegistrations, includeCascadeTraces);
        HtmlComputedStyle? placeholder = element.LocalName.Equals("input", StringComparison.OrdinalIgnoreCase)
            || element.LocalName.Equals("textarea", StringComparison.OrdinalIgnoreCase)
            ? ComputePseudoElementStyle(element, originatingStyle, candidateRules, HtmlPseudoElementKind.Placeholder, budget, containerContexts, environment, customPropertyRegistrations, includeCascadeTraces)
            : null;
        if (before == null && after == null && marker == null && footnoteCall == null && footnoteMarker == null
            && firstLetter == null && firstLine == null && placeholder == null) return;
        pseudoElements[element] = new HtmlPseudoElementStylePair {
            Before = before,
            After = after,
            Marker = marker,
            FootnoteCall = footnoteCall,
            FootnoteMarker = footnoteMarker,
            FirstLetter = firstLetter,
            FirstLine = firstLine,
            Placeholder = placeholder
        };
    }

    private static HtmlComputedStyle? ComputePseudoElementStyle(
        IElement element,
        HtmlComputedStyle originatingStyle,
        IReadOnlyList<StyleRule> candidateRules,
        HtmlPseudoElementKind kind,
        HtmlCssProcessingBudget budget,
        IReadOnlyList<ContainerQueryContext> containerContexts,
        MediaEnvironment environment,
        IReadOnlyDictionary<string, CustomPropertyRegistration> customPropertyRegistrations,
        bool includeCascadeTraces) {
        List<StyleRule>? matchedRules = null;
        foreach (StyleRule rule in candidateRules) {
            budget.RecordSelectorEvaluation();
            if (!AreContainerConditionsApplicable(rule.ContainerConditions, containerContexts, environment)
                || !TryParsePseudoElementSelector(rule.Selector, out string hostSelector, out HtmlPseudoElementKind ruleKind)
                || ruleKind != kind
                || !MatchesSelector(element, rule, budget)) {
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
                    valueAlreadyValidated: true,
                    declarationOrder: declaration.Value.DeclarationOrder,
                    customPropertyRegistrations: customPropertyRegistrations,
                    source: OfficeIMO.Html.Css.HtmlCssCascadeSourceKind.StyleRule,
                    selector: rule.Selector, layerName: rule.LayerName);
            }
        }

        Dictionary<string, string> resolvedProperties = ResolveComputedProperties(properties, originatingStyle.Properties,
            out HashSet<string> inheritedProperties, out HashSet<string> resetProperties,
            out HashSet<string> originRevertedProperties,
            out HashSet<string> specifiedProperties,
            out Dictionary<string, HtmlCssCascadePriority> cascadePriorities,
            out Dictionary<string, OfficeIMO.Html.Css.HtmlCssCascadeTrace>? cascadeTraces,
            includeCascadeTraces,
            customPropertyRegistrations);
        return HtmlComputedStyle.FromOwnedCollections(
            resolvedProperties, inheritedProperties, resetProperties, originRevertedProperties,
            specifiedProperties, cascadePriorities, cascadeTraces);
    }

}
