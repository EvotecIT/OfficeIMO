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
                AddStyleRules(rule, rules, environment, budget, 1);
            }
            IReadOnlyDictionary<int, int> rawRuleClosures = BuildRawRuleClosures(css, budget);
            AddRawRetainedStyleRules(css, 0, css.Length, rawRuleClosures, rules, environment, budget);
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
        HtmlCssProcessingBudget budget,
        int depth) {
        budget.RecordNestingDepth(depth);
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
            AddStyleRules(childRule, rules, environment, budget, depth + 1);
        }
    }

    private static void AddStyleRule(
        AngleSharp.Css.Dom.ICssStyleRule styleRule,
        ICollection<StyleRule> rules,
        HtmlCssProcessingBudget budget) {
        string[] selectors = SplitSelectorList(styleRule.SelectorText)
            .Select(selector => selector.Trim())
            .Where(selector => selector.Length > 0)
            .ToArray();
        foreach (string _ in selectors) budget.RecordRule(styleRule.Style.Length);

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
        AddRetainedUnknownDeclarations(styleRule.CssText, declarations);

        foreach (string selector in selectors) {
            if (declarations.Count > 0) {
                rules.Add(new StyleRule(selector, CalculateSpecificity(selector), rules.Count, declarations));
            }
        }
    }

    private static void AddRawRetainedStyleRules(
        string css,
        int start,
        int end,
        IReadOnlyDictionary<int, int> closures,
        ICollection<StyleRule> rules,
        MediaEnvironment environment,
        HtmlCssProcessingBudget budget) {
        int cursor = start;
        while (cursor < end) {
            SkipRawWhitespaceAndComments(css, ref cursor, end);
            if (cursor >= end) break;
            int preludeStart = cursor;
            int open = FindRawRuleOpen(css, cursor, end);
            if (open < 0) {
                if (open != -1) {
                    cursor = ~open + 1;
                    continue;
                }
                break;
            }
            if (!closures.TryGetValue(open, out int close) || close >= end) break;
            string prelude = css.Substring(preludeStart, open - preludeStart).Trim();
            if (prelude.StartsWith("@media", StringComparison.OrdinalIgnoreCase)) {
                string condition = prelude.Substring(6).Trim();
                if (IsApplicableMedia(condition, environment)) {
                    AddRawRetainedStyleRules(css, open + 1, close, closures, rules, environment, budget);
                }
            } else if (prelude.StartsWith("@supports", StringComparison.OrdinalIgnoreCase)) {
                string condition = prelude.Substring(9).Trim();
                if (IsApplicableSupports(condition)) {
                    AddRawRetainedStyleRules(css, open + 1, close, closures, rules, environment, budget);
                }
            } else if (!prelude.StartsWith("@", StringComparison.Ordinal)) {
                AddRawRetainedStyleRule(prelude, css.Substring(open + 1, close - open - 1), rules, budget);
            }
            cursor = close + 1;
        }
    }

    private static void AddRawRetainedStyleRule(
        string selectorText,
        string body,
        ICollection<StyleRule> rules,
        HtmlCssProcessingBudget budget) {
        var declarations = new Dictionary<string, StyleDeclaration>(StringComparer.OrdinalIgnoreCase);
        foreach (string declaration in SplitCssDeclarations(StripCssCommentsOutsideStrings(body))) {
            int separator = declaration.IndexOf(':');
            if (separator <= 0) continue;
            string propertyName = declaration.Substring(0, separator).Trim();
            if (!string.Equals(propertyName, "string-set", StringComparison.OrdinalIgnoreCase)) continue;
            string value = declaration.Substring(separator + 1).Trim();
            value = StripTrailingImportant(value, out bool important);
            if (value.Length > 0 && IsSupportedDeclarationValue(propertyName, value)) {
                declarations[propertyName] = new StyleDeclaration(value, important);
            }
        }
        if (declarations.Count == 0) return;
        string[] selectors = SplitSelectorList(selectorText)
            .Select(selector => selector.Trim())
            .Where(selector => selector.Length > 0)
            .ToArray();
        foreach (string _ in selectors) budget.RecordRule(declarations.Count);
        foreach (string selector in selectors) {
            rules.Add(new StyleRule(selector, CalculateSpecificity(selector), rules.Count, declarations));
        }
    }

    private static int FindRawRuleOpen(string css, int start, int end) {
        char quote = '\0';
        for (int index = start; index < end; index++) {
            char current = css[index];
            if (quote != '\0') {
                if (current == quote && !IsEscaped(css, index)) quote = '\0';
                continue;
            }
            if (current == '\'' || current == '"') quote = current;
            else if (current == '/' && index + 1 < end && css[index + 1] == '*') {
                index = css.IndexOf("*/", index + 2, StringComparison.Ordinal);
                if (index < 0 || index >= end) return -1;
                index++;
            } else if (current == '{') return index;
            else if (current == ';') return ~index;
        }
        return -1;
    }

    private static IReadOnlyDictionary<int, int> BuildRawRuleClosures(
        string css,
        HtmlCssProcessingBudget budget) {
        var closures = new Dictionary<int, int>();
        var opens = new Stack<int>();
        char quote = '\0';
        for (int index = 0; index < css.Length; index++) {
            char current = css[index];
            if (quote != '\0') {
                if (current == quote && !IsEscaped(css, index)) quote = '\0';
                continue;
            }
            if (current == '\'' || current == '"') quote = current;
            else if (current == '/' && index + 1 < css.Length && css[index + 1] == '*') {
                index = css.IndexOf("*/", index + 2, StringComparison.Ordinal);
                if (index < 0) break;
                index++;
            } else if (current == '{') {
                opens.Push(index);
                budget.RecordNestingDepth(opens.Count);
            } else if (current == '}' && opens.Count > 0) {
                closures[opens.Pop()] = index;
            }
        }
        return closures;
    }

    private static void SkipRawWhitespaceAndComments(string css, ref int cursor, int end) {
        while (cursor < end) {
            if (char.IsWhiteSpace(css[cursor])) {
                cursor++;
                continue;
            }
            if (css[cursor] == '/' && cursor + 1 < end && css[cursor + 1] == '*') {
                int close = css.IndexOf("*/", cursor + 2, StringComparison.Ordinal);
                cursor = close < 0 || close >= end ? end : close + 2;
                continue;
            }
            break;
        }
    }

    private static void AddRetainedUnknownDeclarations(
        string cssText,
        IDictionary<string, StyleDeclaration> declarations) {
        int open = cssText.IndexOf('{');
        int close = cssText.LastIndexOf('}');
        if (open < 0 || close <= open) return;
        string body = StripCssCommentsOutsideStrings(cssText.Substring(open + 1, close - open - 1));
        foreach (string declaration in SplitCssDeclarations(body)) {
            int separator = declaration.IndexOf(':');
            if (separator <= 0) continue;
            string propertyName = declaration.Substring(0, separator).Trim();
            if (declarations.ContainsKey(propertyName)
                || (!SupportedProperties.Contains(propertyName) && !propertyName.StartsWith("--", StringComparison.Ordinal))) continue;
            string value = declaration.Substring(separator + 1).Trim();
            value = StripTrailingImportant(value, out bool important);
            if (value.Length > 0 && IsSupportedDeclarationValue(propertyName, value)) {
                declarations[propertyName] = new StyleDeclaration(value, important);
            }
        }
    }
}
