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
        "color",
        "direction",
        "font-family",
        "font-size",
        "font-style",
        "font-weight",
        "line-height",
        "orphans",
        "page",
        "text-align",
        "text-transform",
        "visibility",
        "widows",
        "white-space"
    };
    private static readonly HashSet<string> SupportedProperties = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
        "background",
        "background-color",
        "background-image",
        "background-position",
        "background-repeat",
        "background-size",
        "align-content",
        "align-items",
        "align-self",
        "bottom",
        "border",
        "border-bottom",
        "border-bottom-color",
        "border-bottom-style",
        "border-bottom-width",
        "border-collapse",
        "border-color",
        "border-left",
        "border-left-color",
        "border-left-style",
        "border-left-width",
        "border-right",
        "border-right-color",
        "border-right-style",
        "border-right-width",
        "border-style",
        "border-top",
        "border-top-color",
        "border-top-style",
        "border-top-width",
        "border-width",
        "box-sizing",
        "break-after",
        "break-before",
        "break-inside",
        "clear",
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
        "counter-increment",
        "counter-reset",
        "counter-set",
        "cursor",
        "direction",
        "display",
        "font-family",
        "font-size",
        "font-style",
        "font-weight",
        "flex",
        "flex-basis",
        "flex-direction",
        "flex-flow",
        "flex-grow",
        "flex-shrink",
        "flex-wrap",
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
        "left",
        "line-height",
        "list-style",
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
        "opacity",
        "order",
        "orphans",
        "outline-color",
        "overflow",
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
        "place-content",
        "place-items",
        "place-self",
        "right",
        "row-gap",
        "text-align",
        "text-decoration-line",
        "text-transform",
        "top",
        "vertical-align",
        "visibility",
        "white-space",
        "width",
        "widows",
        "word-break",
        "z-index"
    };

    /// <summary>
    /// Computes styles for every element in the supplied document using style tags and inline style attributes.
    /// </summary>
    public static IReadOnlyDictionary<IElement, HtmlComputedStyle> Compute(IHtmlDocument document, HtmlCssMediaContext mediaContext = HtmlCssMediaContext.Screen) {
        return ComputeStyleSet(document, mediaContext, false).Elements;
    }

    internal static HtmlComputedStyleSet ComputeForRendering(IHtmlDocument document, HtmlCssMediaContext mediaContext = HtmlCssMediaContext.Screen) =>
        ComputeStyleSet(document, mediaContext, true);

    private static HtmlComputedStyleSet ComputeStyleSet(IHtmlDocument document, HtmlCssMediaContext mediaContext, bool includePseudoElements) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        IReadOnlyList<StyleRule> rules = ParseStyleRules(document, mediaContext);
        var computed = new Dictionary<IElement, HtmlComputedStyle>();
        var pseudoElements = new Dictionary<IElement, HtmlPseudoElementStylePair>();
        IElement? root = document.DocumentElement ?? document.Body;
        if (root != null) {
            ComputeElement(root, null, rules, computed, pseudoElements, includePseudoElements);
        }

        return new HtmlComputedStyleSet(computed, pseudoElements);
    }

    /// <summary>
    /// Parses raw HTML and computes styles for matching elements.
    /// </summary>
    public static IReadOnlyDictionary<IElement, HtmlComputedStyle> Compute(string html, HtmlCssMediaContext mediaContext = HtmlCssMediaContext.Screen) {
        return Compute(HtmlDocumentParser.ParseDocument(html), mediaContext);
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
        IReadOnlyList<StyleRule> rules,
        IDictionary<IElement, HtmlComputedStyle> computed,
        IDictionary<IElement, HtmlPseudoElementStylePair> pseudoElements,
        bool includePseudoElements) {
        var properties = new Dictionary<string, CascadedProperty>(StringComparer.OrdinalIgnoreCase);
        if (parent != null) {
            foreach (var pair in parent.Properties) {
                if (IsInheritedProperty(pair.Key)) {
                    properties[pair.Key] = new CascadedProperty(pair.Value, false, Specificity.Inherited, -1);
                }
            }
        }

        foreach (StyleRule rule in rules) {
            if (!TryParsePseudoElementSelector(rule.Selector, out _, out _)
                && MatchesSelector(element, rule.Selector)) {
                foreach (var declaration in rule.Declarations) {
                    ApplyDeclaration(properties, parent?.Properties, declaration.Key, declaration.Value.Value, declaration.Value.IsImportant, rule.Specificity, rule.Order);
                }
            }
        }

        ApplyInlineDeclarations(properties, parent?.Properties, element.GetAttribute("style"));
        var style = new HtmlComputedStyle(ResolveComputedProperties(properties, parent?.Properties));
        computed[element] = style;
        if (includePseudoElements) ComputePseudoElementStyles(element, style, rules, pseudoElements);

        foreach (IElement child in element.Children) {
            ComputeElement(child, style, rules, computed, pseudoElements, includePseudoElements);
        }
    }

    private static void ComputePseudoElementStyles(
        IElement element,
        HtmlComputedStyle originatingStyle,
        IReadOnlyList<StyleRule> rules,
        IDictionary<IElement, HtmlPseudoElementStylePair> pseudoElements) {
        HtmlComputedStyle? before = ComputePseudoElementStyle(element, originatingStyle, rules, HtmlPseudoElementKind.Before);
        HtmlComputedStyle? after = ComputePseudoElementStyle(element, originatingStyle, rules, HtmlPseudoElementKind.After);
        if (before == null && after == null) return;
        pseudoElements[element] = new HtmlPseudoElementStylePair { Before = before, After = after };
    }

    private static HtmlComputedStyle? ComputePseudoElementStyle(
        IElement element,
        HtmlComputedStyle originatingStyle,
        IReadOnlyList<StyleRule> rules,
        HtmlPseudoElementKind kind) {
        var properties = new Dictionary<string, CascadedProperty>(StringComparer.OrdinalIgnoreCase);
        foreach (KeyValuePair<string, string> pair in originatingStyle.Properties) {
            if (IsInheritedProperty(pair.Key)) {
                properties[pair.Key] = new CascadedProperty(pair.Value, false, Specificity.Inherited, -1);
            }
        }

        bool matched = false;
        foreach (StyleRule rule in rules) {
            if (!TryParsePseudoElementSelector(rule.Selector, out string hostSelector, out HtmlPseudoElementKind ruleKind)
                || ruleKind != kind
                || !MatchesSelector(element, hostSelector)) {
                continue;
            }

            matched = true;
            foreach (KeyValuePair<string, StyleDeclaration> declaration in rule.Declarations) {
                ApplyDeclaration(
                    properties,
                    originatingStyle.Properties,
                    declaration.Key,
                    declaration.Value.Value,
                    declaration.Value.IsImportant,
                    rule.Specificity,
                    rule.Order);
            }
        }

        return matched
            ? new HtmlComputedStyle(ResolveComputedProperties(properties, originatingStyle.Properties))
            : null;
    }

}
