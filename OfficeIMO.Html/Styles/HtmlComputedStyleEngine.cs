using AngleSharp.Css.Parser;
using AngleSharp.Dom;
using AngleSharp.Html.Dom;

namespace OfficeIMO.Html;

/// <summary>
/// Lightweight computed-style helper for OfficeIMO conversion diagnostics and contract tests.
/// </summary>
public static class HtmlComputedStyleEngine {
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

    /// <summary>
    /// Computes styles for every element in the supplied document using style tags and inline style attributes.
    /// </summary>
    public static IReadOnlyDictionary<IElement, HtmlComputedStyle> Compute(IHtmlDocument document) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        IReadOnlyList<StyleRule> rules = ParseStyleRules(document);
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
    public static IReadOnlyDictionary<IElement, HtmlComputedStyle> Compute(string html) {
        return Compute(HtmlDocumentParser.ParseDocument(html));
    }

    private static void ComputeElement(IElement element, HtmlComputedStyle? parent, IReadOnlyList<StyleRule> rules, IDictionary<IElement, HtmlComputedStyle> computed) {
        var properties = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (parent != null) {
            foreach (var pair in parent.Properties) {
                if (InheritedProperties.Contains(pair.Key)) {
                    properties[pair.Key] = pair.Value;
                }
            }
        }

        foreach (StyleRule rule in rules) {
            if (MatchesSelector(element, rule.Selector)) {
                foreach (var declaration in rule.Declarations) {
                    properties[declaration.Key] = declaration.Value;
                }
            }
        }

        ApplyDeclarations(properties, element.GetAttribute("style"));
        var style = new HtmlComputedStyle(properties);
        computed[element] = style;

        foreach (IElement child in element.Children) {
            ComputeElement(child, style, rules, computed);
        }
    }

    private static IReadOnlyList<StyleRule> ParseStyleRules(IHtmlDocument document) {
        var rules = new List<StyleRule>();
        var parser = new CssParser();
        foreach (IElement styleElement in document.QuerySelectorAll("style")) {
            string css = styleElement.TextContent;
            if (string.IsNullOrWhiteSpace(css)) {
                continue;
            }

            var stylesheet = parser.ParseStyleSheet(css);
            foreach (var rule in stylesheet.Rules) {
                var styleRule = rule as AngleSharp.Css.Dom.ICssStyleRule;
                if (styleRule == null) {
                    continue;
                }

                var declarations = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                for (int i = 0; i < styleRule.Style.Length; i++) {
                    string propertyName = styleRule.Style[i];
                    if (!string.IsNullOrWhiteSpace(propertyName)) {
                        declarations[propertyName] = styleRule.Style.GetPropertyValue(propertyName);
                    }
                }

                foreach (string selector in styleRule.SelectorText.Split(',')) {
                    string trimmedSelector = selector.Trim();
                    if (trimmedSelector.Length > 0 && declarations.Count > 0) {
                        rules.Add(new StyleRule(trimmedSelector, declarations));
                    }
                }
            }
        }

        return rules;
    }

    private static bool MatchesSelector(IElement element, string selector) {
        try {
            return element.Matches(selector);
        } catch (DomException) {
            return MatchesSimpleSelector(element, selector);
        }
    }

    private static bool MatchesSimpleSelector(IElement element, string selector) {
        if (selector.StartsWith(".", StringComparison.Ordinal)) {
            return element.ClassList.Contains(selector.Substring(1));
        }

        if (selector.StartsWith("#", StringComparison.Ordinal)) {
            return string.Equals(element.Id, selector.Substring(1), StringComparison.Ordinal);
        }

        return string.Equals(element.TagName, selector, StringComparison.OrdinalIgnoreCase);
    }

    private static void ApplyDeclarations(IDictionary<string, string> properties, string? styleText) {
        if (string.IsNullOrWhiteSpace(styleText)) {
            return;
        }

        foreach (string declaration in styleText!.Split(';')) {
            int separator = declaration.IndexOf(':');
            if (separator <= 0) {
                continue;
            }

            string name = declaration.Substring(0, separator).Trim();
            string value = declaration.Substring(separator + 1).Trim();
            if (name.Length > 0 && value.Length > 0) {
                properties[name] = value;
            }
        }
    }

    private sealed class StyleRule {
        internal StyleRule(string selector, IDictionary<string, string> declarations) {
            Selector = selector;
            Declarations = new Dictionary<string, string>(declarations, StringComparer.OrdinalIgnoreCase);
        }

        internal string Selector { get; }
        internal IReadOnlyDictionary<string, string> Declarations { get; }
    }
}
