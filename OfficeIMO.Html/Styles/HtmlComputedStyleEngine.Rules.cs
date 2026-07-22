using AngleSharp.Css.Parser;
using AngleSharp.Dom;
using AngleSharp.Html.Dom;

namespace OfficeIMO.Html;

public static partial class HtmlComputedStyleEngine {
    private static IReadOnlyList<StyleRule> ParseStyleRules(
        IHtmlDocument document,
        MediaEnvironment environment,
        HtmlCssProcessingBudget budget) {
        var rules = new List<StyleRule>();
        var parser = new CssParser(new CssParserOptions {
            IsIncludingUnknownDeclarations = true
        });
        foreach (IElement styleElement in document.QuerySelectorAll("style")) {
            if (!IsCssStyleElement(styleElement)) {
                continue;
            }

            if (!IsApplicableMedia(styleElement.GetAttribute("media") ?? string.Empty, environment)) {
                continue;
            }

            string css = styleElement.TextContent;
            if (string.IsNullOrWhiteSpace(css)) {
                continue;
            }

            var stylesheet = parser.ParseStyleSheet(css);
            foreach (var rule in stylesheet.Rules) {
                AddStyleRules(rule, rules, environment, budget);
            }
        }

        return rules;
    }

    private static bool IsEffectivelyHidden(HtmlComputedStyle style) {
        return string.Equals(style.GetValue("display"), "none", StringComparison.OrdinalIgnoreCase)
            || string.Equals(style.GetValue("visibility"), "hidden", StringComparison.OrdinalIgnoreCase)
            || string.Equals(style.GetValue("visibility"), "collapse", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsColorProperty(string propertyName) {
        return propertyName.IndexOf("color", StringComparison.OrdinalIgnoreCase) >= 0;
    }

    private static bool IsCssStyleElement(IElement styleElement) {
        string type = (styleElement.GetAttribute("type") ?? string.Empty).Trim();
        if (type.Length == 0) {
            return true;
        }

        int parameterStart = type.IndexOf(';');
        if (parameterStart >= 0) {
            type = type.Substring(0, parameterStart).Trim();
        }

        return string.Equals(type, "text/css", StringComparison.OrdinalIgnoreCase);
    }

    private static void AddStyleRules(
        AngleSharp.Css.Dom.ICssRule rule,
        ICollection<StyleRule> rules,
        MediaEnvironment environment,
        HtmlCssProcessingBudget budget) {
        var styleRule = rule as AngleSharp.Css.Dom.ICssStyleRule;
        if (styleRule != null) {
            AddStyleRule(styleRule, rules, budget);
            return;
        }

        var mediaRule = rule as AngleSharp.Css.Dom.ICssMediaRule;
        if (mediaRule != null && !IsApplicableMedia(mediaRule.ConditionText, environment)) {
            return;
        }

        if (IsSupportsRule(rule) && !IsApplicableSupports(GetConditionText(rule))) {
            return;
        }

        var groupingRule = rule as AngleSharp.Css.Dom.ICssGroupingRule;
        if (groupingRule == null) {
            return;
        }

        foreach (var childRule in groupingRule.Rules) {
            AddStyleRules(childRule, rules, environment, budget);
        }
    }

    private static void AddStyleRule(
        AngleSharp.Css.Dom.ICssStyleRule styleRule,
        ICollection<StyleRule> rules,
        HtmlCssProcessingBudget budget) {
        var declarations = new Dictionary<string, StyleDeclaration>(StringComparer.OrdinalIgnoreCase);
        for (int i = 0; i < styleRule.Style.Length; i++) {
            string propertyName = styleRule.Style[i];
            if (!string.IsNullOrWhiteSpace(propertyName)
                && (SupportedProperties.Contains(propertyName) || propertyName.StartsWith("--", StringComparison.Ordinal))) {
                declarations[propertyName] = new StyleDeclaration(
                    styleRule.Style.GetPropertyValue(propertyName),
                    string.Equals(styleRule.Style.GetPropertyPriority(propertyName), "important", StringComparison.OrdinalIgnoreCase));
            }
        }

        // AngleSharp can retain a var()-backed shorthand while enumerating only empty
        // expanded longhands. Query supported properties directly so the cascade keeps
        // the authored shorthand for custom-property resolution.
        foreach (string propertyName in SupportedProperties) {
            if (declarations.ContainsKey(propertyName)) continue;
            string propertyValue = styleRule.Style.GetPropertyValue(propertyName);
            if (string.IsNullOrWhiteSpace(propertyValue)) continue;
            declarations[propertyName] = new StyleDeclaration(
                propertyValue,
                string.Equals(styleRule.Style.GetPropertyPriority(propertyName), "important", StringComparison.OrdinalIgnoreCase));
        }

        foreach (string selector in SplitSelectorList(styleRule.SelectorText)) {
            string trimmedSelector = selector.Trim();
            if (trimmedSelector.Length > 0 && declarations.Count > 0) {
                budget.RecordRule(declarations.Count);
                rules.Add(new StyleRule(trimmedSelector, CalculateSpecificity(trimmedSelector), rules.Count, declarations));
            }
        }
    }
}
