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
        "text-align",
        "text-transform",
        "visibility",
        "white-space"
    };
    private static readonly HashSet<string> SupportedProperties = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
        "background",
        "background-color",
        "background-image",
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
        "color",
        "cursor",
        "direction",
        "display",
        "font-family",
        "font-size",
        "font-style",
        "font-weight",
        "height",
        "line-height",
        "list-style",
        "list-style-type",
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
        "outline-color",
        "overflow",
        "overflow-wrap",
        "page-break-after",
        "page-break-before",
        "page-break-inside",
        "padding",
        "padding-bottom",
        "padding-left",
        "padding-right",
        "padding-top",
        "text-align",
        "text-decoration-line",
        "text-transform",
        "vertical-align",
        "visibility",
        "white-space",
        "width",
        "word-break"
    };

    /// <summary>
    /// Computes styles for every element in the supplied document using style tags and inline style attributes.
    /// </summary>
    public static IReadOnlyDictionary<IElement, HtmlComputedStyle> Compute(IHtmlDocument document, HtmlCssMediaContext mediaContext = HtmlCssMediaContext.Screen) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        IReadOnlyList<StyleRule> rules = ParseStyleRules(document, mediaContext);
        var computed = new Dictionary<IElement, HtmlComputedStyle>();
        IElement? root = document.DocumentElement ?? document.Body;
        if (root != null) {
            ComputeElement(root, null, rules, computed);
        }

        return computed;
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

    private static void ComputeElement(IElement element, HtmlComputedStyle? parent, IReadOnlyList<StyleRule> rules, IDictionary<IElement, HtmlComputedStyle> computed) {
        var properties = new Dictionary<string, CascadedProperty>(StringComparer.OrdinalIgnoreCase);
        if (parent != null) {
            foreach (var pair in parent.Properties) {
                if (IsInheritedProperty(pair.Key)) {
                    properties[pair.Key] = new CascadedProperty(pair.Value, false, Specificity.Inherited, -1);
                }
            }
        }

        foreach (StyleRule rule in rules) {
            if (MatchesSelector(element, rule.Selector)) {
                foreach (var declaration in rule.Declarations) {
                    ApplyDeclaration(properties, parent?.Properties, declaration.Key, declaration.Value.Value, declaration.Value.IsImportant, rule.Specificity, rule.Order);
                }
            }
        }

        ApplyInlineDeclarations(properties, parent?.Properties, element.GetAttribute("style"));
        var style = new HtmlComputedStyle(ResolveComputedProperties(properties, parent?.Properties));
        computed[element] = style;

        foreach (IElement child in element.Children) {
            ComputeElement(child, style, rules, computed);
        }
    }

}
