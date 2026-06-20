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
        var properties = new Dictionary<string, CascadedProperty>(StringComparer.OrdinalIgnoreCase);
        if (parent != null) {
            foreach (var pair in parent.Properties) {
                if (InheritedProperties.Contains(pair.Key)) {
                    properties[pair.Key] = new CascadedProperty(pair.Value, false, -1, -1);
                }
            }
        }

        foreach (StyleRule rule in rules) {
            if (MatchesSelector(element, rule.Selector)) {
                foreach (var declaration in rule.Declarations) {
                    ApplyDeclaration(properties, declaration.Key, declaration.Value.Value, declaration.Value.IsImportant, rule.Specificity, rule.Order);
                }
            }
        }

        ApplyInlineDeclarations(properties, element.GetAttribute("style"));
        var style = new HtmlComputedStyle(properties.ToDictionary(pair => pair.Key, pair => pair.Value.Value, StringComparer.OrdinalIgnoreCase));
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

                var declarations = new Dictionary<string, StyleDeclaration>(StringComparer.OrdinalIgnoreCase);
                for (int i = 0; i < styleRule.Style.Length; i++) {
                    string propertyName = styleRule.Style[i];
                    if (!string.IsNullOrWhiteSpace(propertyName)) {
                        declarations[propertyName] = new StyleDeclaration(
                            styleRule.Style.GetPropertyValue(propertyName),
                            string.Equals(styleRule.Style.GetPropertyPriority(propertyName), "important", StringComparison.OrdinalIgnoreCase));
                    }
                }

                foreach (string selector in SplitSelectorList(styleRule.SelectorText)) {
                    string trimmedSelector = selector.Trim();
                    if (trimmedSelector.Length > 0 && declarations.Count > 0) {
                        rules.Add(new StyleRule(trimmedSelector, CalculateSpecificity(trimmedSelector), rules.Count, declarations));
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

    private static void ApplyInlineDeclarations(IDictionary<string, CascadedProperty> properties, string? styleText) {
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
            bool isImportant = EndsWithImportant(value);
            if (isImportant) {
                value = value.Substring(0, value.Length - 10).TrimEnd();
            }

            if (name.Length > 0 && value.Length > 0) {
                ApplyDeclaration(properties, name, value, isImportant, 1000000, int.MaxValue);
            }
        }
    }

    private static void ApplyDeclaration(IDictionary<string, CascadedProperty> properties, string name, string value, bool isImportant, int specificity, int order) {
        if (string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(value)) {
            return;
        }

        CascadedProperty? existing;
        if (properties.TryGetValue(name, out existing) && existing != null && !ShouldReplace(existing, isImportant, specificity, order)) {
            return;
        }

        properties[name] = new CascadedProperty(value, isImportant, specificity, order);
    }

    private static bool ShouldReplace(CascadedProperty existing, bool isImportant, int specificity, int order) {
        if (existing.IsImportant != isImportant) {
            return isImportant;
        }

        if (existing.Specificity != specificity) {
            return specificity >= existing.Specificity;
        }

        return order >= existing.Order;
    }

    private static bool EndsWithImportant(string value) {
        return value.EndsWith("!important", StringComparison.OrdinalIgnoreCase);
    }

    private static IEnumerable<string> SplitSelectorList(string selectorText) {
        int depth = 0;
        char quote = '\0';
        int start = 0;
        for (int i = 0; i < selectorText.Length; i++) {
            char current = selectorText[i];
            if (quote != '\0') {
                if (current == quote && !IsEscaped(selectorText, i)) {
                    quote = '\0';
                }

                continue;
            }

            if (current == '"' || current == '\'') {
                quote = current;
                continue;
            }

            if (current == '(' || current == '[') {
                depth++;
                continue;
            }

            if (current == ')' || current == ']') {
                if (depth > 0) {
                    depth--;
                }

                continue;
            }

            if (current == ',' && depth == 0) {
                yield return selectorText.Substring(start, i - start);
                start = i + 1;
            }
        }

        yield return selectorText.Substring(start);
    }

    private static bool IsEscaped(string text, int index) {
        int slashCount = 0;
        for (int i = index - 1; i >= 0 && text[i] == '\\'; i--) {
            slashCount++;
        }

        return slashCount % 2 == 1;
    }

    private static int CalculateSpecificity(string selector) {
        int ids = 0;
        int classesAttributesAndPseudoClasses = 0;
        int elements = 0;
        bool inAttribute = false;
        for (int i = 0; i < selector.Length; i++) {
            char current = selector[i];
            if (current == '[') {
                inAttribute = true;
                classesAttributesAndPseudoClasses++;
                continue;
            }

            if (current == ']') {
                inAttribute = false;
                continue;
            }

            if (inAttribute) {
                continue;
            }

            if (current == '#') {
                ids++;
                i = SkipIdentifier(selector, i + 1);
            } else if (current == '.') {
                classesAttributesAndPseudoClasses++;
                i = SkipIdentifier(selector, i + 1);
            } else if (current == ':') {
                if (i + 1 < selector.Length && selector[i + 1] == ':') {
                    elements++;
                    i = SkipIdentifier(selector, i + 2);
                } else {
                    classesAttributesAndPseudoClasses++;
                    i = SkipIdentifier(selector, i + 1);
                }
            } else if (IsElementStart(selector, i)) {
                elements++;
                i = SkipIdentifier(selector, i);
            }
        }

        return (ids * 10000) + (classesAttributesAndPseudoClasses * 100) + elements;
    }

    private static int SkipIdentifier(string selector, int index) {
        int current = index;
        while (current < selector.Length && IsIdentifierCharacter(selector[current])) {
            current++;
        }

        return current - 1;
    }

    private static bool IsElementStart(string selector, int index) {
        char current = selector[index];
        if (!char.IsLetter(current) && current != '*') {
            return false;
        }

        if (index > 0) {
            char previous = selector[index - 1];
            if (previous == '#' || previous == '.' || previous == '-' || previous == '_' || previous == ':') {
                return false;
            }
        }

        return true;
    }

    private static bool IsIdentifierCharacter(char value) {
        return char.IsLetterOrDigit(value) || value == '-' || value == '_';
    }

    private sealed class StyleDeclaration {
        internal StyleDeclaration(string value, bool isImportant) {
            Value = value;
            IsImportant = isImportant;
        }

        internal string Value { get; }
        internal bool IsImportant { get; }
    }

    private sealed class CascadedProperty {
        internal CascadedProperty(string value, bool isImportant, int specificity, int order) {
            Value = value;
            IsImportant = isImportant;
            Specificity = specificity;
            Order = order;
        }

        internal string Value { get; }
        internal bool IsImportant { get; }
        internal int Specificity { get; }
        internal int Order { get; }
    }

    private sealed class StyleRule {
        internal StyleRule(string selector, int specificity, int order, IDictionary<string, StyleDeclaration> declarations) {
            Selector = selector;
            Specificity = specificity;
            Order = order;
            Declarations = new Dictionary<string, StyleDeclaration>(declarations, StringComparer.OrdinalIgnoreCase);
        }

        internal string Selector { get; }
        internal int Specificity { get; }
        internal int Order { get; }
        internal IReadOnlyDictionary<string, StyleDeclaration> Declarations { get; }
    }
}
