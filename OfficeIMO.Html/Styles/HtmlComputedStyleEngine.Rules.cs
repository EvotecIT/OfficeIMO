using AngleSharp.Css.Parser;
using AngleSharp.Dom;
using AngleSharp.Html.Dom;

namespace OfficeIMO.Html;

public static partial class HtmlComputedStyleEngine {
    private const string RevertLayerSentinel = "var(--officeimo-internal-revert-layer)";

    private static IReadOnlyList<StyleRule> ParseStyleRules(
        IHtmlDocument document,
        MediaEnvironment environment,
        HtmlCssProcessingBudget budget) {
        var rules = new List<StyleRule>();
        var layers = new CascadeLayerRegistry();
        // Raw recovery supplements declarations AngleSharp cannot retain. Match each
        // parsed author rule once so recovery neither duplicates it nor charges it twice.
        var parsedRuleMatches = new Dictionary<string, int>(StringComparer.Ordinal);
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

            IReadOnlyDictionary<int, int> rawRuleClosures = HtmlCssRuleBlockScanner.Scan(css, budget);
            string parseCss = ExpandNestedConditionalRules(css);
            parseCss = PreserveManagedGradientFunctions(PreserveRevertLayerDeclarations(parseCss));
            var stylesheet = parser.ParseStyleSheet(parseCss);
            foreach (var rule in stylesheet.Rules) {
                AddStyleRules(rule, rules, parsedRuleMatches, environment, budget, layers, 1, null, null, null);
            }
            AddRawRetainedStyleRules(css, 0, css.Length, rawRuleClosures, rules, parsedRuleMatches, environment, budget);
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
        IDictionary<string, int> parsedRuleMatches,
        MediaEnvironment environment,
        HtmlCssProcessingBudget budget,
        CascadeLayerRegistry layers,
        int depth,
        string? currentLayer,
        IReadOnlyList<string>? parentSelectors,
        IReadOnlyList<ContainerRuleCondition>? containerConditions) {
        budget.RecordNestingDepth(depth);
        var layerRule = rule as AngleSharp.Css.Dom.ICssLayerRule;
        if (layerRule != null) {
            if (layerRule.IsStatement) {
                layers.RegisterStatement(layerRule.Name, currentLayer);
                return;
            }
            (string layerName, _) = layers.RegisterBlock(layerRule.Name, currentLayer);
            foreach (var childRule in layerRule.Rules) {
                AddStyleRules(childRule, rules, parsedRuleMatches, environment, budget, layers, depth + 1, layerName, parentSelectors, containerConditions);
            }
            return;
        }

        var containerRule = rule as AngleSharp.Css.Dom.ICssContainerRule;
        if (containerRule != null) {
            string containerName = containerRule.ContainerName?.Trim() ?? string.Empty;
            string containerQuery = containerRule.ContainerQuery?.Trim() ?? string.Empty;
            if (string.Equals(containerName, "style", StringComparison.OrdinalIgnoreCase)
                && containerQuery.StartsWith("(", StringComparison.Ordinal)) {
                containerName = string.Empty;
                containerQuery = "style" + containerQuery;
            }
            var nestedConditions = new List<ContainerRuleCondition>(containerConditions ?? Array.Empty<ContainerRuleCondition>()) {
                new ContainerRuleCondition(containerName, containerQuery)
            };
            foreach (var childRule in containerRule.Rules) {
                AddStyleRules(childRule, rules, parsedRuleMatches, environment, budget, layers, depth + 1, currentLayer, parentSelectors, nestedConditions);
            }
            return;
        }

        var styleRule = rule as AngleSharp.Css.Dom.ICssStyleRule;
        if (styleRule != null) {
            IReadOnlyList<string> resolvedSelectors = ResolveNestedSelectors(styleRule.SelectorText ?? string.Empty, parentSelectors);
            AddStyleRule(styleRule, resolvedSelectors, rules, parsedRuleMatches, budget, currentLayer == null ? null : layers.GetOrder(currentLayer), containerConditions);
            foreach (var childRule in styleRule.Rules) {
                AddStyleRules(childRule, rules, parsedRuleMatches, environment, budget, layers, depth + 1, currentLayer, resolvedSelectors, containerConditions);
            }
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
            AddStyleRules(childRule, rules, parsedRuleMatches, environment, budget, layers, depth + 1, currentLayer, parentSelectors, containerConditions);
        }
    }

    private static void AddStyleRule(
        AngleSharp.Css.Dom.ICssStyleRule styleRule,
        IReadOnlyList<string> resolvedSelectors,
        ICollection<StyleRule> rules,
        IDictionary<string, int> parsedRuleMatches,
        HtmlCssProcessingBudget budget,
        CascadeLayerOrder? layerOrder,
        IReadOnlyList<ContainerRuleCondition>? containerConditions) {
        if (resolvedSelectors.Count == 0) return;
        string[] selectors = resolvedSelectors.ToArray();
        foreach (string selector in selectors) {
            budget.RecordRule(styleRule.Style.Length);
            RecordParsedRule(parsedRuleMatches, ParsedRuleKey(selector));
        }

        var declarations = new Dictionary<string, StyleDeclaration>(StringComparer.OrdinalIgnoreCase);
        for (int i = 0; i < styleRule.Style.Length; i++) {
            string propertyName = styleRule.Style[i];
            if (!string.IsNullOrWhiteSpace(propertyName)
                && (SupportedProperties.Contains(propertyName) || propertyName.StartsWith("--", StringComparison.Ordinal))) {
                declarations[propertyName] = new StyleDeclaration(
                    RestoreProtectedDeclarationValue(styleRule.Style.GetPropertyValue(propertyName)),
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
                RestoreProtectedDeclarationValue(propertyValue),
                string.Equals(styleRule.Style.GetPropertyPriority(propertyName), "important", StringComparison.OrdinalIgnoreCase));
        }
        AddRetainedUnknownDeclarations(styleRule.CssText, declarations);

        foreach (string selector in selectors) {
            if (declarations.Count > 0) {
                rules.Add(new StyleRule(selector, CalculateSpecificity(selector), rules.Count, declarations, layerOrder, containerConditions));
                if (declarations.ContainsKey("string-set")) {
                    RecordParsedRule(parsedRuleMatches, ParsedRetainedRuleKey(selector));
                }
            }
        }
    }

    private static IReadOnlyList<string> ResolveNestedSelectors(string selectorText, IReadOnlyList<string>? parentSelectors) {
        string[] children = SplitSelectorList(selectorText)
            .Select(selector => selector.Trim())
            .Where(selector => selector.Length > 0)
            .ToArray();
        if (parentSelectors == null || parentSelectors.Count == 0) return children;
        string parent = parentSelectors.Count == 1
            ? parentSelectors[0]
            : ":is(" + string.Join(",", parentSelectors) + ")";
        var resolved = new List<string>(children.Length);
        foreach (string child in children) {
            string nested = ReplaceNestingSelectorTokens(child, parent, out bool replaced);
            resolved.Add(replaced
                ? nested
                : parent + " " + child);
        }
        return resolved.AsReadOnly();
    }

    private static string ReplaceNestingSelectorTokens(string selector, string parent, out bool replaced) {
        replaced = false;
        var result = new System.Text.StringBuilder(selector.Length + parent.Length);
        char quote = '\0';
        int attributeDepth = 0;
        for (int index = 0; index < selector.Length; index++) {
            char current = selector[index];
            if (current == '\\') {
                result.Append(current);
                if (index + 1 < selector.Length) result.Append(selector[++index]);
                continue;
            }
            if (quote != '\0') {
                result.Append(current);
                if (current == quote) quote = '\0';
                continue;
            }
            if (current == '/' && index + 1 < selector.Length && selector[index + 1] == '*') {
                int commentEnd = selector.IndexOf("*/", index + 2, StringComparison.Ordinal);
                if (commentEnd < 0) {
                    result.Append(selector, index, selector.Length - index);
                    break;
                }
                result.Append(selector, index, commentEnd + 2 - index);
                index = commentEnd + 1;
                continue;
            }
            if (current is '\'' or '"') {
                quote = current;
                result.Append(current);
                continue;
            }
            if (current == '[') {
                attributeDepth++;
                result.Append(current);
                continue;
            }
            if (current == ']' && attributeDepth > 0) {
                attributeDepth--;
                result.Append(current);
                continue;
            }
            if (current == '&' && attributeDepth == 0) {
                result.Append(parent);
                replaced = true;
                continue;
            }
            result.Append(current);
        }
        return result.ToString();
    }

    private static string RestoreRevertLayerKeyword(string value) =>
        string.Equals(value.Trim(), RevertLayerSentinel, StringComparison.OrdinalIgnoreCase)
            ? "revert-layer"
            : value;

    private static string RestoreProtectedDeclarationValue(string value) =>
        RestoreManagedGradientFunctions(RestoreRevertLayerKeyword(value));

    private static string PreserveRevertLayerDeclarations(string css) {
        const string keyword = "revert-layer";
        var result = new System.Text.StringBuilder(css.Length);
        char quote = '\0';
        for (int index = 0; index < css.Length;) {
            char current = css[index];
            if (quote != '\0') {
                result.Append(current);
                if (current == quote && !IsEscaped(css, index)) quote = '\0';
                index++;
                continue;
            }
            if (current == '\'' || current == '"') {
                quote = current;
                result.Append(current);
                index++;
                continue;
            }
            if (current == '/' && index + 1 < css.Length && css[index + 1] == '*') {
                int close = css.IndexOf("*/", index + 2, StringComparison.Ordinal);
                int end = close < 0 ? css.Length : close + 2;
                result.Append(css, index, end - index);
                index = end;
                continue;
            }
            if (index + keyword.Length <= css.Length
                && string.Compare(css, index, keyword, 0, keyword.Length, StringComparison.OrdinalIgnoreCase) == 0
                && IsStandaloneDeclarationValue(css, index, keyword.Length)) {
                result.Append(RevertLayerSentinel);
                index += keyword.Length;
                continue;
            }
            result.Append(current);
            index++;
        }
        return result.ToString();
    }

    private static bool IsStandaloneDeclarationValue(string css, int start, int length) {
        int before = start - 1;
        while (before >= 0 && char.IsWhiteSpace(css[before])) before--;
        if (before < 0 || css[before] != ':') return false;
        int after = start + length;
        while (after < css.Length && char.IsWhiteSpace(css[after])) after++;
        const string important = "!important";
        if (after + important.Length <= css.Length
            && string.Compare(css, after, important, 0, important.Length, StringComparison.OrdinalIgnoreCase) == 0) {
            after += important.Length;
            while (after < css.Length && char.IsWhiteSpace(css[after])) after++;
        }
        return after < css.Length && (css[after] == ';' || css[after] == '}');
    }

    private static string ExpandNestedConditionalRules(string css) {
        if (string.IsNullOrEmpty(css) || css.IndexOf('@') < 0) return css;
        var output = new System.Text.StringBuilder(css.Length);
        int cursor = 0;
        while (TryFindNextTopLevelBlock(css, cursor, out int preludeStart, out int open, out int close)) {
            output.Append(css, cursor, preludeStart - cursor);
            string prelude = css.Substring(preludeStart, open - preludeStart).Trim();
            string body = css.Substring(open + 1, close - open - 1);
            if (IsConditionalGroupingPrelude(prelude)) {
                output.Append(prelude).Append('{').Append(ExpandNestedConditionalRules(body)).Append('}');
            } else if (!prelude.StartsWith("@", StringComparison.Ordinal)) {
                AppendNestedStyleRuleExpansion(output, prelude, body);
            } else {
                output.Append(prelude).Append('{').Append(body).Append('}');
            }
            cursor = close + 1;
        }
        output.Append(css, cursor, css.Length - cursor);
        return output.ToString();
    }

    private static void AppendNestedStyleRuleExpansion(
        System.Text.StringBuilder output,
        string selector,
        string body) {
        int cursor = 0;
        while (TryFindNestedConditionalBlock(body, cursor, out int start, out int open, out int close)) {
            AppendStyleRuleSegment(output, selector, body, cursor, start - cursor);
            string prelude = body.Substring(start, open - start).Trim();
            string nestedBody = body.Substring(open + 1, close - open - 1);
            output.Append(prelude).Append('{')
                .Append(ExpandNestedConditionalRules(selector + "{" + nestedBody + "}"))
                .Append('}');
            cursor = close + 1;
        }
        AppendStyleRuleSegment(output, selector, body, cursor, body.Length - cursor);
    }

    private static void AppendStyleRuleSegment(
        System.Text.StringBuilder output,
        string selector,
        string body,
        int start,
        int length) {
        if (length <= 0) return;
        string segment = body.Substring(start, length);
        if (string.IsNullOrWhiteSpace(segment)) return;
        output.Append(selector).Append('{').Append(segment).Append('}');
    }

    private static bool TryFindNextTopLevelBlock(
        string css,
        int start,
        out int preludeStart,
        out int open,
        out int close) {
        preludeStart = open = close = -1;
        int candidateStart = start;
        char quote = '\0';
        for (int index = start; index < css.Length; index++) {
            char current = css[index];
            if (quote != '\0') {
                if (current == quote && !IsEscaped(css, index)) quote = '\0';
                continue;
            }
            if (current == '\'' || current == '"') {
                quote = current;
                continue;
            }
            if (current == '/' && index + 1 < css.Length && css[index + 1] == '*') {
                int commentClose = css.IndexOf("*/", index + 2, StringComparison.Ordinal);
                index = commentClose < 0 ? css.Length : commentClose + 1;
                continue;
            }
            if (current == ';') {
                candidateStart = index + 1;
                continue;
            }
            if (current != '{') continue;
            int matching = FindMatchingCssBrace(css, index);
            if (matching < 0) return false;
            preludeStart = candidateStart;
            while (preludeStart < index && char.IsWhiteSpace(css[preludeStart])) preludeStart++;
            open = index;
            close = matching;
            return preludeStart < open;
        }
        return false;
    }

    private static bool TryFindNestedConditionalBlock(
        string body,
        int start,
        out int blockStart,
        out int open,
        out int close) {
        blockStart = open = close = -1;
        int depth = 0;
        char quote = '\0';
        for (int index = start; index < body.Length; index++) {
            char current = body[index];
            if (quote != '\0') {
                if (current == quote && !IsEscaped(body, index)) quote = '\0';
                continue;
            }
            if (current == '\'' || current == '"') {
                quote = current;
                continue;
            }
            if (current == '/' && index + 1 < body.Length && body[index + 1] == '*') {
                int commentClose = body.IndexOf("*/", index + 2, StringComparison.Ordinal);
                index = commentClose < 0 ? body.Length : commentClose + 1;
                continue;
            }
            if (current == '{') {
                depth++;
                continue;
            }
            if (current == '}' && depth > 0) {
                depth--;
                continue;
            }
            if (depth != 0 || current != '@') continue;
            int keywordEnd = index + 1;
            while (keywordEnd < body.Length && (char.IsLetter(body[keywordEnd]) || body[keywordEnd] == '-')) keywordEnd++;
            string keyword = body.Substring(index, keywordEnd - index);
            if (!IsConditionalGroupingPrelude(keyword)) continue;
            int candidateOpen = FindConditionalBlockOpen(body, keywordEnd);
            if (candidateOpen < 0) continue;
            int candidateClose = FindMatchingCssBrace(body, candidateOpen);
            if (candidateClose < 0) return false;
            blockStart = index;
            open = candidateOpen;
            close = candidateClose;
            return true;
        }
        return false;
    }

    private static int FindConditionalBlockOpen(string text, int start) {
        int parentheses = 0;
        char quote = '\0';
        for (int index = start; index < text.Length; index++) {
            char current = text[index];
            if (quote != '\0') {
                if (current == quote && !IsEscaped(text, index)) quote = '\0';
                continue;
            }
            if (current == '\'' || current == '"') quote = current;
            else if (current == '(') parentheses++;
            else if (current == ')' && parentheses > 0) parentheses--;
            else if (parentheses == 0 && current == ';') return -1;
            else if (parentheses == 0 && current == '{') return index;
        }
        return -1;
    }

    private static int FindMatchingCssBrace(string text, int open) {
        int depth = 0;
        char quote = '\0';
        for (int index = open; index < text.Length; index++) {
            char current = text[index];
            if (quote != '\0') {
                if (current == quote && !IsEscaped(text, index)) quote = '\0';
                continue;
            }
            if (current == '\'' || current == '"') {
                quote = current;
                continue;
            }
            if (current == '/' && index + 1 < text.Length && text[index + 1] == '*') {
                int commentClose = text.IndexOf("*/", index + 2, StringComparison.Ordinal);
                index = commentClose < 0 ? text.Length : commentClose + 1;
                continue;
            }
            if (current == '{') depth++;
            else if (current == '}' && --depth == 0) return index;
        }
        return -1;
    }

    private static bool IsConditionalGroupingPrelude(string prelude) {
        string normalized = prelude.TrimStart();
        return StartsWithAtRuleName(normalized, "@media")
            || StartsWithAtRuleName(normalized, "@supports")
            || StartsWithAtRuleName(normalized, "@layer")
            || StartsWithAtRuleName(normalized, "@container");
    }

    private static bool StartsWithAtRuleName(string value, string name) =>
        value.StartsWith(name, StringComparison.OrdinalIgnoreCase)
        && (value.Length == name.Length || char.IsWhiteSpace(value[name.Length]) || value[name.Length] == '(');

    private static void AddRawRetainedStyleRules(
        string css,
        int start,
        int end,
        IReadOnlyDictionary<int, int> closures,
        ICollection<StyleRule> rules,
        IDictionary<string, int> parsedRuleMatches,
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
                    AddRawRetainedStyleRules(css, open + 1, close, closures, rules, parsedRuleMatches, environment, budget);
                }
            } else if (prelude.StartsWith("@supports", StringComparison.OrdinalIgnoreCase)) {
                string condition = prelude.Substring(9).Trim();
                if (IsApplicableSupports(condition)) {
                    AddRawRetainedStyleRules(css, open + 1, close, closures, rules, parsedRuleMatches, environment, budget);
                }
            } else if (!prelude.StartsWith("@", StringComparison.Ordinal)) {
                AddRawRetainedStyleRule(prelude, css.Substring(open + 1, close - open - 1), rules, parsedRuleMatches, budget);
            }
            cursor = close + 1;
        }
    }

    private static void AddRawRetainedStyleRule(
        string selectorText,
        string body,
        ICollection<StyleRule> rules,
        IDictionary<string, int> parsedRuleMatches,
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
                SetDeclarationInSourceOrder(declarations, propertyName, value, important);
            }
        }
        if (declarations.Count == 0) return;
        string[] selectors = SplitSelectorList(selectorText)
            .Select(selector => selector.Trim())
            .Where(selector => selector.Length > 0)
            .ToArray();
        foreach (string selector in selectors) {
            bool parsedRule = TryConsumeParsedRule(parsedRuleMatches, ParsedRuleKey(selector));
            if (TryConsumeParsedRule(parsedRuleMatches, ParsedRetainedRuleKey(selector))) continue;
            if (!parsedRule) budget.RecordRule(declarations.Count);
            rules.Add(new StyleRule(selector, CalculateSpecificity(selector), rules.Count, declarations));
        }
    }

    private static void RecordParsedRule(
        IDictionary<string, int> parsedRuleMatches,
        string key) {
        parsedRuleMatches.TryGetValue(key, out int count);
        parsedRuleMatches[key] = count + 1;
    }

    private static bool TryConsumeParsedRule(
        IDictionary<string, int> parsedRuleMatches,
        string key) {
        if (!parsedRuleMatches.TryGetValue(key, out int count) || count <= 0) return false;
        if (count == 1) parsedRuleMatches.Remove(key);
        else parsedRuleMatches[key] = count - 1;
        return true;
    }

    private static string ParsedRuleKey(string selector) => "rule\u001f" + selector.Trim();

    private static string ParsedRetainedRuleKey(string selector) => "retained\u001f" + selector.Trim();

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
            } else if (current == '{' && !IsEscaped(css, index)) return index;
            else if (current == ';' && !IsEscaped(css, index)) return ~index;
        }
        return -1;
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
        var parsedProperties = new HashSet<string>(declarations.Keys, StringComparer.OrdinalIgnoreCase);
        foreach (string declaration in SplitCssDeclarations(body)) {
            int separator = declaration.IndexOf(':');
            if (separator <= 0) continue;
            string propertyName = declaration.Substring(0, separator).Trim();
            if (parsedProperties.Contains(propertyName)
                || (!SupportedProperties.Contains(propertyName) && !propertyName.StartsWith("--", StringComparison.Ordinal))) continue;
            string value = declaration.Substring(separator + 1).Trim();
            value = StripTrailingImportant(value, out bool important);
            if (value.Length > 0 && IsSupportedDeclarationValue(propertyName, value)) {
                SetDeclarationInSourceOrder(declarations, propertyName, value, important);
            }
        }
    }

    private static void SetDeclarationInSourceOrder(
        IDictionary<string, StyleDeclaration> declarations,
        string propertyName,
        string value,
        bool important) {
        if (declarations.TryGetValue(propertyName, out StyleDeclaration? existing)
            && existing.IsImportant
            && !important) {
            return;
        }
        declarations[propertyName] = new StyleDeclaration(value, important);
    }
}
