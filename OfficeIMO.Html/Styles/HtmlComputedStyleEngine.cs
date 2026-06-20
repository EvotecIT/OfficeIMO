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
    private static readonly HashSet<string> SupportedProperties = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
        "background",
        "background-image",
        "border",
        "border-color",
        "color",
        "cursor",
        "display",
        "font-family",
        "font-size",
        "font-style",
        "font-weight",
        "line-height",
        "list-style",
        "outline-color",
        "padding",
        "text-align",
        "text-decoration-line",
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
        var style = new HtmlComputedStyle(properties
            .Where(pair => pair.Value.HasValue)
            .ToDictionary(pair => pair.Key, pair => pair.Value.Value, StringComparer.OrdinalIgnoreCase));
        computed[element] = style;

        foreach (IElement child in element.Children) {
            ComputeElement(child, style, rules, computed);
        }
    }

    private static IReadOnlyList<StyleRule> ParseStyleRules(IHtmlDocument document) {
        var rules = new List<StyleRule>();
        var parser = new CssParser();
        foreach (IElement styleElement in document.QuerySelectorAll("style")) {
            if (!IsCssStyleElement(styleElement)) {
                continue;
            }

            if (!IsApplicableMedia(styleElement.GetAttribute("media") ?? string.Empty)) {
                continue;
            }

            string css = styleElement.TextContent;
            if (string.IsNullOrWhiteSpace(css)) {
                continue;
            }

            var stylesheet = parser.ParseStyleSheet(css);
            foreach (var rule in stylesheet.Rules) {
                AddStyleRules(rule, rules);
            }
        }

        return rules;
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

    private static void AddStyleRules(AngleSharp.Css.Dom.ICssRule rule, ICollection<StyleRule> rules) {
        var styleRule = rule as AngleSharp.Css.Dom.ICssStyleRule;
        if (styleRule != null) {
            AddStyleRule(styleRule, rules);
            return;
        }

        var mediaRule = rule as AngleSharp.Css.Dom.ICssMediaRule;
        if (mediaRule != null && !IsApplicableMedia(mediaRule.ConditionText)) {
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
            AddStyleRules(childRule, rules);
        }
    }

    private static void AddStyleRule(AngleSharp.Css.Dom.ICssStyleRule styleRule, ICollection<StyleRule> rules) {
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

    private static bool IsApplicableMedia(string mediaText) {
        if (string.IsNullOrWhiteSpace(mediaText)) {
            return true;
        }

        foreach (string query in SplitSelectorList(mediaText)) {
            string normalized = query.Trim();
            if (normalized.StartsWith("not ", StringComparison.OrdinalIgnoreCase)) {
                string negated = normalized.Substring(4).Trim();
                if (ContainsMediaType(negated, "screen") || ContainsMediaType(negated, "all")) {
                    continue;
                }

                if (ContainsMediaType(negated, "print")) {
                    return true;
                }

                if (HasMediaFeatureConstraint(negated)) {
                    continue;
                }

                continue;
            }

            if (HasMediaFeatureConstraint(normalized)) {
                continue;
            }

            if (ContainsMediaType(normalized, "all") || ContainsMediaType(normalized, "screen")) {
                return true;
            }
        }

        return false;
    }

    private static bool ContainsMediaType(string mediaQuery, string mediaType) {
        foreach (string token in mediaQuery.Split(new[] { ' ', '\t', '\r', '\n', '\f' }, StringSplitOptions.RemoveEmptyEntries)) {
            if (string.Equals(token.Trim(), mediaType, StringComparison.OrdinalIgnoreCase)) {
                return true;
            }
        }

        return false;
    }

    private static bool HasMediaFeatureConstraint(string mediaQuery) {
        return mediaQuery.IndexOf("(", StringComparison.Ordinal) >= 0
            || mediaQuery.IndexOf(":", StringComparison.Ordinal) >= 0;
    }

    private static bool IsSupportsRule(AngleSharp.Css.Dom.ICssRule rule) {
        string name = rule.GetType().Name;
        string? fullName = rule.GetType().FullName;
        return name.IndexOf("Supports", StringComparison.OrdinalIgnoreCase) >= 0
            || (fullName != null && fullName.IndexOf("Supports", StringComparison.OrdinalIgnoreCase) >= 0);
    }

    private static string GetConditionText(AngleSharp.Css.Dom.ICssRule rule) {
        var property = rule.GetType().GetProperty("ConditionText");
        object? value = property?.GetValue(rule, null);
        return value as string ?? string.Empty;
    }

    private static bool IsApplicableSupports(string conditionText) {
        if (string.IsNullOrWhiteSpace(conditionText)) {
            return true;
        }

        return EvaluateSupportsCondition(conditionText.Trim());
    }

    private static bool EvaluateSupportsCondition(string conditionText) {
        string normalized = conditionText.Trim();
        if (normalized.Length == 0) {
            return true;
        }

        if (StartsWithLogicalNot(normalized)) {
            return !EvaluateSupportsCondition(normalized.Substring(3).TrimStart());
        }

        List<string> orParts = SplitTopLevelLogical(normalized, "or").ToList();
        if (orParts.Count > 1) {
            return orParts.Any(EvaluateSupportsCondition);
        }

        List<string> andParts = SplitTopLevelLogical(normalized, "and").ToList();
        if (andParts.Count > 1) {
            return andParts.All(EvaluateSupportsCondition);
        }

        if (normalized[0] == '(') {
            int close = FindMatchingParenthesis(normalized, 0);
            if (close == normalized.Length - 1) {
                return EvaluateSupportsCondition(normalized.Substring(1, normalized.Length - 2));
            }
        }

        int separator = normalized.IndexOf(':');
        if (separator <= 0) {
            return false;
        }

        string propertyName = normalized.Substring(0, separator).Trim();
        string value = normalized.Substring(separator + 1).Trim();
        return IsSupportedDeclarationValue(propertyName, value);
    }

    private static bool IsSupportedDeclarationValue(string propertyName, string value) {
        if (!SupportedProperties.Contains(propertyName) || string.IsNullOrWhiteSpace(value)) {
            return false;
        }

        string normalized = value.Trim().Trim('\'', '"').ToLowerInvariant();
        switch (propertyName.ToLowerInvariant()) {
            case "display":
                return IsKnownKeyword(normalized, "block", "inline", "inline-block", "none", "flex", "inline-flex", "grid", "inline-grid", "table", "table-row", "table-cell", "list-item", "contents", "flow-root");
            case "visibility":
                return IsKnownKeyword(normalized, "visible", "hidden", "collapse");
            case "text-transform":
                return IsKnownKeyword(normalized, "none", "uppercase", "lowercase", "capitalize", "full-width", "full-size-kana");
            case "text-decoration-line":
                return normalized.Split(new[] { ' ', '\t', '\r', '\n', '\f' }, StringSplitOptions.RemoveEmptyEntries)
                    .All(token => IsKnownKeyword(token, "none", "underline", "overline", "line-through", "blink"));
            case "font-style":
                return normalized == "normal" || normalized == "italic" || normalized.StartsWith("oblique", StringComparison.Ordinal);
            case "font-weight":
                int weight;
                return IsKnownKeyword(normalized, "normal", "bold", "bolder", "lighter")
                    || (int.TryParse(normalized, out weight) && weight >= 1 && weight <= 1000);
            case "text-align":
                return IsKnownKeyword(normalized, "left", "right", "center", "justify", "start", "end", "match-parent");
            case "direction":
                return IsKnownKeyword(normalized, "ltr", "rtl");
            case "white-space":
                return IsKnownKeyword(normalized, "normal", "nowrap", "pre", "pre-wrap", "pre-line", "break-spaces");
            default:
                return !normalized.StartsWith("not-a-real", StringComparison.Ordinal);
        }
    }

    private static bool IsKnownKeyword(string value, params string[] keywords) {
        foreach (string keyword in keywords) {
            if (string.Equals(value, keyword, StringComparison.OrdinalIgnoreCase)) {
                return true;
            }
        }

        return false;
    }

    private static bool StartsWithLogicalNot(string conditionText) {
        return conditionText.Length > 3
            && conditionText.StartsWith("not", StringComparison.OrdinalIgnoreCase)
            && char.IsWhiteSpace(conditionText[3]);
    }

    private static IEnumerable<string> SplitTopLevelLogical(string conditionText, string logicalOperator) {
        int depth = 0;
        char quote = '\0';
        int start = 0;
        for (int i = 0; i < conditionText.Length; i++) {
            char current = conditionText[i];
            if (quote != '\0') {
                if (current == quote && !IsEscaped(conditionText, i)) {
                    quote = '\0';
                }

                continue;
            }

            if (current == '"' || current == '\'') {
                quote = current;
                continue;
            }

            if (current == '(') {
                depth++;
                continue;
            }

            if (current == ')') {
                if (depth > 0) {
                    depth--;
                }

                continue;
            }

            if (depth == 0 && IsLogicalOperatorAt(conditionText, i, logicalOperator)) {
                yield return conditionText.Substring(start, i - start).Trim();
                i += logicalOperator.Length - 1;
                start = i + 1;
            }
        }

        yield return conditionText.Substring(start).Trim();
    }

    private static bool IsLogicalOperatorAt(string conditionText, int index, string logicalOperator) {
        if (index < 0 || index + logicalOperator.Length > conditionText.Length) {
            return false;
        }

        if (string.Compare(conditionText, index, logicalOperator, 0, logicalOperator.Length, StringComparison.OrdinalIgnoreCase) != 0) {
            return false;
        }

        bool hasLeftBoundary = index == 0 || char.IsWhiteSpace(conditionText[index - 1]);
        int after = index + logicalOperator.Length;
        bool hasRightBoundary = after >= conditionText.Length || char.IsWhiteSpace(conditionText[after]);
        return hasLeftBoundary && hasRightBoundary;
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

    private static void ApplyInlineDeclarations(IDictionary<string, CascadedProperty> properties, IReadOnlyDictionary<string, string>? parentProperties, string? styleText) {
        if (string.IsNullOrWhiteSpace(styleText)) {
            return;
        }

        foreach (string declaration in SplitCssDeclarations(StripCssCommentsOutsideStrings(styleText!))) {
            int separator = declaration.IndexOf(':');
            if (separator <= 0) {
                continue;
            }

            string name = declaration.Substring(0, separator).Trim();
            string value = declaration.Substring(separator + 1).Trim();
            bool isImportant;
            value = StripTrailingImportant(value, out isImportant);

            if (name.Length > 0 && value.Length > 0) {
                ApplyDeclaration(properties, parentProperties, name, value, isImportant, Specificity.Inline, int.MaxValue);
            }
        }
    }

    private static void ApplyDeclaration(IDictionary<string, CascadedProperty> properties, IReadOnlyDictionary<string, string>? parentProperties, string name, string value, bool isImportant, Specificity specificity, int order) {
        if (string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(value)) {
            return;
        }

        var resolved = ResolveCssWideKeyword(name, value, parentProperties);
        if (!resolved.HasValue) {
            CascadedProperty? resetExisting;
            if (properties.TryGetValue(name, out resetExisting) && resetExisting != null && !ShouldReplace(resetExisting, isImportant, specificity, order)) {
                return;
            }

            properties[name] = CascadedProperty.Clear(isImportant, specificity, order);
            return;
        }

        CascadedProperty? existing;
        if (properties.TryGetValue(name, out existing) && existing != null && !ShouldReplace(existing, isImportant, specificity, order)) {
            return;
        }

        properties[name] = new CascadedProperty(resolved.Value, isImportant, specificity, order);
    }

    private static CssKeywordResolution ResolveCssWideKeyword(string name, string value, IReadOnlyDictionary<string, string>? parentProperties) {
        string trimmed = value.Trim();
        if (string.Equals(trimmed, "inherit", StringComparison.OrdinalIgnoreCase)
            || (string.Equals(trimmed, "unset", StringComparison.OrdinalIgnoreCase) && InheritedProperties.Contains(name))) {
            string? inheritedValue;
            return parentProperties != null && parentProperties.TryGetValue(name, out inheritedValue) && !string.IsNullOrWhiteSpace(inheritedValue)
                ? CssKeywordResolution.ForValue(inheritedValue)
                : CssKeywordResolution.Clear;
        }

        if (string.Equals(trimmed, "initial", StringComparison.OrdinalIgnoreCase)
            || string.Equals(trimmed, "revert", StringComparison.OrdinalIgnoreCase)
            || string.Equals(trimmed, "revert-layer", StringComparison.OrdinalIgnoreCase)) {
            return string.Equals(name, "visibility", StringComparison.OrdinalIgnoreCase)
                ? CssKeywordResolution.ForValue("visible")
                : CssKeywordResolution.Clear;
        }

        if (string.Equals(trimmed, "unset", StringComparison.OrdinalIgnoreCase)) {
            return CssKeywordResolution.Clear;
        }

        return CssKeywordResolution.ForValue(value);
    }

    private static bool ShouldReplace(CascadedProperty existing, bool isImportant, Specificity specificity, int order) {
        if (existing.IsImportant != isImportant) {
            return isImportant;
        }

        int specificityComparison = specificity.CompareTo(existing.Specificity);
        if (specificityComparison != 0) {
            return specificityComparison > 0;
        }

        return order >= existing.Order;
    }

    private static string StripTrailingImportant(string value, out bool isImportant) {
        isImportant = false;
        if (string.IsNullOrWhiteSpace(value)) {
            return value;
        }

        string trimmed = value.TrimEnd();
        const string ImportantKeyword = "important";
        int importantStart = trimmed.Length - ImportantKeyword.Length;
        if (importantStart < 0 || !string.Equals(trimmed.Substring(importantStart), ImportantKeyword, StringComparison.OrdinalIgnoreCase)) {
            return value;
        }

        int bangIndex = importantStart - 1;
        while (bangIndex >= 0 && char.IsWhiteSpace(trimmed[bangIndex])) {
            bangIndex--;
        }

        if (bangIndex < 0 || trimmed[bangIndex] != '!') {
            return value;
        }

        if (IsInsideCssString(trimmed, bangIndex) || IsInsideCssComment(trimmed, bangIndex)) {
            return value;
        }

        isImportant = true;
        return trimmed.Substring(0, bangIndex).TrimEnd();
    }

    private static bool IsInsideCssComment(string text, int index) {
        int open = text.LastIndexOf("/*", Math.Max(0, index), StringComparison.Ordinal);
        if (open < 0) {
            return false;
        }

        int close = text.LastIndexOf("*/", Math.Max(0, index), StringComparison.Ordinal);
        return close < open;
    }

    private static string StripCssCommentsOutsideStrings(string css) {
        var result = new System.Text.StringBuilder(css.Length);
        char quote = '\0';
        for (int i = 0; i < css.Length; i++) {
            char current = css[i];
            if (quote != '\0') {
                result.Append(current);
                if (current == quote && !IsEscaped(css, i)) {
                    quote = '\0';
                }

                continue;
            }

            if (current == '"' || current == '\'') {
                quote = current;
                result.Append(current);
                continue;
            }

            if (current == '/' && i + 1 < css.Length && css[i + 1] == '*') {
                i += 2;
                while (i + 1 < css.Length && !(css[i] == '*' && css[i + 1] == '/')) {
                    i++;
                }

                if (i + 1 < css.Length) {
                    i++;
                }

                result.Append(' ');
                continue;
            }

            result.Append(current);
        }

        return result.ToString();
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

    private static IEnumerable<string> SplitCssDeclarations(string styleText) {
        int depth = 0;
        char quote = '\0';
        int start = 0;
        for (int i = 0; i < styleText.Length; i++) {
            char current = styleText[i];
            if (quote != '\0') {
                if (current == quote && !IsEscaped(styleText, i)) {
                    quote = '\0';
                }

                continue;
            }

            if (current == '"' || current == '\'') {
                quote = current;
                continue;
            }

            if (current == '(') {
                depth++;
                continue;
            }

            if (current == ')') {
                if (depth > 0) {
                    depth--;
                }

                continue;
            }

            if (current == ';' && depth == 0) {
                yield return styleText.Substring(start, i - start);
                start = i + 1;
            }
        }

        yield return styleText.Substring(start);
    }

    private static bool IsInsideCssString(string text, int index) {
        char quote = '\0';
        for (int i = 0; i < index && i < text.Length; i++) {
            char current = text[i];
            if (quote != '\0') {
                if (current == quote && !IsEscaped(text, i)) {
                    quote = '\0';
                }

                continue;
            }

            if (current == '"' || current == '\'') {
                quote = current;
            }
        }

        return quote != '\0';
    }

    private static bool IsEscaped(string text, int index) {
        int slashCount = 0;
        for (int i = index - 1; i >= 0 && text[i] == '\\'; i--) {
            slashCount++;
        }

        return slashCount % 2 == 1;
    }

    private static Specificity CalculateSpecificity(string selector) {
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
                    int nameStart = i + 1;
                    int nameEnd = SkipIdentifier(selector, nameStart);
                    string pseudoName = selector.Substring(nameStart, nameEnd - nameStart + 1);
                    if (nameEnd + 1 < selector.Length && selector[nameEnd + 1] == '(') {
                        int close = FindMatchingParenthesis(selector, nameEnd + 1);
                        if (close > nameEnd + 1) {
                            string argument = selector.Substring(nameEnd + 2, close - nameEnd - 2);
                            if (string.Equals(pseudoName, "where", StringComparison.OrdinalIgnoreCase)) {
                                i = close;
                                continue;
                            }

                            if (string.Equals(pseudoName, "is", StringComparison.OrdinalIgnoreCase)
                                || string.Equals(pseudoName, "not", StringComparison.OrdinalIgnoreCase)
                                || string.Equals(pseudoName, "has", StringComparison.OrdinalIgnoreCase)) {
                                Specificity argumentSpecificity = MaxSpecificity(argument);
                                ids += argumentSpecificity.Ids;
                                classesAttributesAndPseudoClasses += argumentSpecificity.ClassesAttributesAndPseudoClasses;
                                elements += argumentSpecificity.Elements;
                                i = close;
                                continue;
                            }

                            classesAttributesAndPseudoClasses++;
                            i = close;
                            continue;
                        }
                    }

                    classesAttributesAndPseudoClasses++;
                    i = nameEnd;
                }
            } else if (IsElementStart(selector, i)) {
                elements++;
                i = SkipIdentifier(selector, i);
            }
        }

        return new Specificity(ids, classesAttributesAndPseudoClasses, elements);
    }

    private static Specificity MaxSpecificity(string selectorList) {
        var max = new Specificity(0, 0, 0);
        foreach (string selector in SplitSelectorList(selectorList)) {
            Specificity specificity = CalculateSpecificity(selector);
            if (specificity.CompareTo(max) > 0) {
                max = specificity;
            }
        }

        return max;
    }

    private static int FindMatchingParenthesis(string text, int openIndex) {
        int depth = 0;
        char quote = '\0';
        for (int i = openIndex; i < text.Length; i++) {
            char current = text[i];
            if (quote != '\0') {
                if (current == quote && !IsEscaped(text, i)) {
                    quote = '\0';
                }

                continue;
            }

            if (current == '"' || current == '\'') {
                quote = current;
                continue;
            }

            if (current == '(') {
                depth++;
            } else if (current == ')') {
                depth--;
                if (depth == 0) {
                    return i;
                }
            }
        }

        return -1;
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
        if (!char.IsLetter(current)) {
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
        internal CascadedProperty(string value, bool isImportant, Specificity specificity, int order) {
            Value = value;
            HasValue = true;
            IsImportant = isImportant;
            Specificity = specificity;
            Order = order;
        }

        private CascadedProperty(bool isImportant, Specificity specificity, int order) {
            Value = string.Empty;
            HasValue = false;
            IsImportant = isImportant;
            Specificity = specificity;
            Order = order;
        }

        internal static CascadedProperty Clear(bool isImportant, Specificity specificity, int order) {
            return new CascadedProperty(isImportant, specificity, order);
        }

        internal string Value { get; }
        internal bool HasValue { get; }
        internal bool IsImportant { get; }
        internal Specificity Specificity { get; }
        internal int Order { get; }
    }

    private readonly struct CssKeywordResolution {
        private CssKeywordResolution(bool hasValue, string value) {
            HasValue = hasValue;
            Value = value;
        }

        internal static CssKeywordResolution Clear => new CssKeywordResolution(false, string.Empty);
        internal static CssKeywordResolution ForValue(string value) => new CssKeywordResolution(true, value);

        internal bool HasValue { get; }
        internal string Value { get; }
    }

    private sealed class Specificity {
        internal Specificity(int ids, int classesAttributesAndPseudoClasses, int elements) {
            Ids = ids;
            ClassesAttributesAndPseudoClasses = classesAttributesAndPseudoClasses;
            Elements = elements;
        }

        internal int Ids { get; }
        internal int ClassesAttributesAndPseudoClasses { get; }
        internal int Elements { get; }
        internal static Specificity Inherited { get; } = new Specificity(-1, -1, -1);
        internal static Specificity Inline { get; } = new Specificity(int.MaxValue, int.MaxValue, int.MaxValue);

        internal int CompareTo(Specificity other) {
            if (Ids != other.Ids) {
                return Ids.CompareTo(other.Ids);
            }

            if (ClassesAttributesAndPseudoClasses != other.ClassesAttributesAndPseudoClasses) {
                return ClassesAttributesAndPseudoClasses.CompareTo(other.ClassesAttributesAndPseudoClasses);
            }

            return Elements.CompareTo(other.Elements);
        }
    }

    private sealed class StyleRule {
        internal StyleRule(string selector, Specificity specificity, int order, IDictionary<string, StyleDeclaration> declarations) {
            Selector = selector;
            Specificity = specificity;
            Order = order;
            Declarations = new Dictionary<string, StyleDeclaration>(declarations, StringComparer.OrdinalIgnoreCase);
        }

        internal string Selector { get; }
        internal Specificity Specificity { get; }
        internal int Order { get; }
        internal IReadOnlyDictionary<string, StyleDeclaration> Declarations { get; }
    }
}
