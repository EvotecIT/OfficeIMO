namespace OfficeIMO.Html;

public static partial class HtmlComputedStyleEngine {
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

    /// <summary>
    /// Evaluates whether a CSS supports condition is active for the OfficeIMO CSS subset.
    /// </summary>
    public static bool IsApplicableSupports(string conditionText) {
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
        return IsSupportedSupportsConditionValue(propertyName, value);
    }

    private static bool IsSupportedSupportsConditionValue(string propertyName, string value) {
        string normalized = value.Trim().Trim('\'', '"').ToLowerInvariant();
        if (string.Equals(propertyName, "float", StringComparison.OrdinalIgnoreCase)) {
            return IsKnownKeyword(normalized, "none", "left", "right", "inline-start", "inline-end");
        }
        if (string.Equals(propertyName, "clear", StringComparison.OrdinalIgnoreCase)) {
            return IsKnownKeyword(normalized, "none", "left", "right", "both", "inline-start", "inline-end");
        }
        return IsSupportedDeclarationValue(propertyName, value);
    }

    private static bool IsSupportedDeclarationValue(string propertyName, string value) {
        if (propertyName.StartsWith("--", StringComparison.Ordinal) && !string.IsNullOrWhiteSpace(value)) {
            return true;
        }

        if (!SupportedProperties.Contains(propertyName) || string.IsNullOrWhiteSpace(value)) {
            return false;
        }

        if (HtmlCssCustomPropertyResolver.ContainsVarFunction(value)) {
            return true;
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

    private static bool IsInheritedProperty(string propertyName) =>
        InheritedProperties.Contains(propertyName) || propertyName.StartsWith("--", StringComparison.Ordinal);

    private static IDictionary<string, string> ResolveComputedProperties(
        IReadOnlyDictionary<string, CascadedProperty> properties,
        IReadOnlyDictionary<string, string>? parentProperties) {
        var raw = properties
            .Where(pair => pair.Value.HasValue)
            .ToDictionary(pair => pair.Key, pair => pair.Value.Value, StringComparer.OrdinalIgnoreCase);
        var resolved = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (KeyValuePair<string, string> pair in raw) {
            if (pair.Key.StartsWith("--", StringComparison.Ordinal)) {
                resolved[pair.Key] = pair.Value;
                continue;
            }

            bool success = HtmlCssCustomPropertyResolver.TryResolve(
                pair.Value,
                name => raw.TryGetValue(name, out string? local)
                    ? local
                    : parentProperties != null && parentProperties.TryGetValue(name, out string? inherited) ? inherited : null,
                out string value);
            if (success && IsSupportedDeclarationValue(pair.Key, value)) {
                resolved[pair.Key] = value;
            }
        }

        return resolved;
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

}
