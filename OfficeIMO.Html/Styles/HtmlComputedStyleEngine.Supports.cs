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
        if (string.Equals(propertyName, "overflow", StringComparison.OrdinalIgnoreCase)) {
            string[] values = normalized.Split(new[] { ' ', '\t', '\r', '\n', '\f' }, StringSplitOptions.RemoveEmptyEntries);
            return values.Length >= 1 && values.Length <= 2
                && values.All(item => IsKnownKeyword(item, "visible", "hidden", "clip", "auto", "scroll"));
        }
        if (string.Equals(propertyName, "overflow-x", StringComparison.OrdinalIgnoreCase)
            || string.Equals(propertyName, "overflow-y", StringComparison.OrdinalIgnoreCase)) {
            return IsKnownKeyword(normalized, "visible", "hidden", "clip", "auto", "scroll");
        }
        if (string.Equals(propertyName, "column-count", StringComparison.OrdinalIgnoreCase)) {
            return normalized == "auto" || int.TryParse(normalized, out int count) && count > 0;
        }
        if (string.Equals(propertyName, "column-fill", StringComparison.OrdinalIgnoreCase)) {
            return IsKnownKeyword(normalized, "auto", "balance");
        }
        if (string.Equals(propertyName, "column-span", StringComparison.OrdinalIgnoreCase)) {
            return IsKnownKeyword(normalized, "none", "all");
        }
        if (string.Equals(propertyName, "column-width", StringComparison.OrdinalIgnoreCase)) {
            return normalized == "auto" || IsPositiveCssLength(normalized);
        }
        if (string.Equals(propertyName, "columns", StringComparison.OrdinalIgnoreCase)) {
            IReadOnlyList<string> values = HtmlRenderCssValues.SplitWhitespace(normalized);
            if (values.Count == 0 || values.Count > 2) return false;
            bool hasCount = false;
            bool hasWidth = false;
            foreach (string item in values) {
                if (item == "auto") continue;
                if (!hasCount && int.TryParse(item, out int count) && count > 0) {
                    hasCount = true;
                    continue;
                }
                if (!hasWidth && IsPositiveCssLength(item)) {
                    hasWidth = true;
                    continue;
                }
                return false;
            }
            return true;
        }
        if (string.Equals(propertyName, "column-rule-style", StringComparison.OrdinalIgnoreCase)) {
            return IsKnownKeyword(normalized, "none", "hidden", "solid", "dashed", "dotted", "double");
        }
        if (string.Equals(propertyName, "column-rule-width", StringComparison.OrdinalIgnoreCase)) {
            return IsKnownKeyword(normalized, "thin", "medium", "thick") || IsNonNegativeCssLength(normalized);
        }
        if (string.Equals(propertyName, "column-rule-color", StringComparison.OrdinalIgnoreCase)) {
            return normalized == "currentcolor" || HtmlRenderCssValues.TryColor(normalized, out _);
        }
        if (string.Equals(propertyName, "column-rule", StringComparison.OrdinalIgnoreCase)) {
            IReadOnlyList<string> values = HtmlRenderCssValues.SplitWhitespace(normalized);
            if (values.Count == 0 || values.Count > 3) return false;
            bool hasWidth = false;
            bool hasStyle = false;
            bool hasColor = false;
            foreach (string item in values) {
                if (!hasWidth && (IsKnownKeyword(item, "thin", "medium", "thick") || IsNonNegativeCssLength(item))) {
                    hasWidth = true;
                    continue;
                }
                if (!hasStyle && IsKnownKeyword(item, "none", "hidden", "solid", "dashed", "dotted", "double")) {
                    hasStyle = true;
                    continue;
                }
                if (!hasColor && (item == "currentcolor" || HtmlRenderCssValues.TryColor(item, out _))) {
                    hasColor = true;
                    continue;
                }
                return false;
            }
            return true;
        }
        if (string.Equals(propertyName, "opacity", StringComparison.OrdinalIgnoreCase)) {
            string number = normalized.EndsWith("%", StringComparison.Ordinal)
                ? normalized.Substring(0, normalized.Length - 1)
                : normalized;
            return double.TryParse(number, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double opacity)
                && !double.IsNaN(opacity)
                && !double.IsInfinity(opacity);
        }
        if (string.Equals(propertyName, "transform", StringComparison.OrdinalIgnoreCase)) {
            return HtmlCssTransformParser.IsSupportedTransformSyntax(normalized);
        }
        if (string.Equals(propertyName, "transform-origin", StringComparison.OrdinalIgnoreCase)) {
            return HtmlCssTransformParser.IsSupportedOriginSyntax(normalized);
        }
        return IsSupportedDeclarationValue(propertyName, value);
    }

    private static bool IsPositiveCssLength(string value) {
        int unitStart = 0;
        while (unitStart < value.Length && (char.IsDigit(value[unitStart]) || value[unitStart] == '.' || value[unitStart] == '+' || value[unitStart] == '-')) unitStart++;
        if (unitStart == 0 || unitStart == value.Length) return false;
        if (!double.TryParse(value.Substring(0, unitStart), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double length)
            || length <= 0D || double.IsNaN(length) || double.IsInfinity(length)) return false;
        string unit = value.Substring(unitStart);
        return IsKnownKeyword(unit, "px", "pt", "pc", "in", "cm", "mm", "q", "em", "rem");
    }

    private static bool IsNonNegativeCssLength(string value) {
        if (value == "0") return true;
        int unitStart = 0;
        while (unitStart < value.Length && (char.IsDigit(value[unitStart]) || value[unitStart] == '.' || value[unitStart] == '+' || value[unitStart] == '-')) unitStart++;
        if (unitStart == 0 || unitStart == value.Length) return false;
        if (!double.TryParse(value.Substring(0, unitStart), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double length)
            || length < 0D || double.IsNaN(length) || double.IsInfinity(length)) return false;
        return IsKnownKeyword(value.Substring(unitStart), "px", "pt", "pc", "in", "cm", "mm", "q", "em", "rem");
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
