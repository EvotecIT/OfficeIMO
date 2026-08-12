using System.Globalization;

namespace OfficeIMO.Html;

public static partial class HtmlComputedStyleEngine {
    private static bool AreContainerConditionsApplicable(
        IReadOnlyList<ContainerRuleCondition> conditions,
        IReadOnlyList<ContainerQueryContext> contexts,
        MediaEnvironment environment) {
        foreach (ContainerRuleCondition condition in conditions) {
            ContainerQueryContext? context = FindContainerContext(condition, contexts);
            if (context == null || !EvaluateContainerCondition(condition.Condition, context, environment)) return false;
        }
        return true;
    }

    private static ContainerQueryContext? FindContainerContext(
        ContainerRuleCondition condition,
        IReadOnlyList<ContainerQueryContext> contexts) {
        bool needsWidth = ContainsContainerFeature(condition.Condition, "width")
            || ContainsContainerFeature(condition.Condition, "inline-size")
            || ContainsContainerFeature(condition.Condition, "aspect-ratio")
            || ContainsContainerFeature(condition.Condition, "orientation");
        bool needsHeight = ContainsContainerFeature(condition.Condition, "height")
            || ContainsContainerFeature(condition.Condition, "block-size")
            || ContainsContainerFeature(condition.Condition, "aspect-ratio")
            || ContainsContainerFeature(condition.Condition, "orientation");

        for (int index = contexts.Count - 1; index >= 0; index--) {
            ContainerQueryContext context = contexts[index];
            if (condition.Name.Length > 0 && !context.Names.Contains(condition.Name, StringComparer.Ordinal)) continue;
            if (needsWidth && context.Type != "inline-size" && context.Type != "size") continue;
            if (needsHeight && (context.Type != "size" || !context.Height.HasValue)) continue;
            return context;
        }
        return null;
    }

    private static bool ContainsContainerFeature(string condition, string feature) {
        for (int index = 0; index < condition.Length;) {
            if (!char.IsLetter(condition[index]) && condition[index] != '-') {
                index++;
                continue;
            }

            int tokenStart = index;
            while (index < condition.Length && (char.IsLetter(condition[index]) || condition[index] == '-')) index++;
            string token = condition.Substring(tokenStart, index - tokenStart);
            int lookahead = index;
            while (lookahead < condition.Length && char.IsWhiteSpace(condition[lookahead])) lookahead++;
            if (string.Equals(token, "style", StringComparison.OrdinalIgnoreCase)
                && lookahead < condition.Length
                && condition[lookahead] == '(') {
                int close = FindMatchingParenthesis(condition, lookahead);
                index = close < 0 ? condition.Length : close + 1;
                continue;
            }

            if (string.Equals(token, feature, StringComparison.OrdinalIgnoreCase)
                || string.Equals(token, "min-" + feature, StringComparison.OrdinalIgnoreCase)
                || string.Equals(token, "max-" + feature, StringComparison.OrdinalIgnoreCase)) {
                return true;
            }
        }
        return false;
    }

    private static IReadOnlyList<ContainerQueryContext> AddContainerContext(
        HtmlComputedStyle style,
        double width,
        double? height,
        double fontSize,
        double rootFontSize,
        IReadOnlyList<ContainerQueryContext> contexts) {
        string type = ResolveContainerType(style);
        IReadOnlyList<string> names = ResolveContainerNames(style);
        var expanded = new List<ContainerQueryContext>(contexts.Count + 1);
        expanded.AddRange(contexts);
        expanded.Add(new ContainerQueryContext(names, type, width, height, fontSize, rootFontSize, style.Properties));
        return expanded.AsReadOnly();
    }

    private static string ResolveContainerType(HtmlComputedStyle style) {
        string value = style.GetValue("container-type").Trim().ToLowerInvariant();
        if (value == "size" || value == "inline-size") return value;
        string shorthand = style.GetValue("container");
        int slash = shorthand.IndexOf('/');
        if (slash >= 0) {
            value = shorthand.Substring(slash + 1).Trim().ToLowerInvariant();
            if (value == "size" || value == "inline-size") return value;
        }
        return "normal";
    }

    private static IReadOnlyList<string> ResolveContainerNames(HtmlComputedStyle style) {
        string value = style.GetValue("container-name").Trim();
        if (value.Length == 0) {
            string shorthand = style.GetValue("container");
            int slash = shorthand.IndexOf('/');
            value = (slash >= 0 ? shorthand.Substring(0, slash) : shorthand).Trim();
        }
        if (value.Length == 0 || string.Equals(value, "none", StringComparison.OrdinalIgnoreCase)) return Array.Empty<string>();
        return value.Split(new[] { ' ', '\t', '\r', '\n', '\f' }, StringSplitOptions.RemoveEmptyEntries).ToList().AsReadOnly();
    }

    private static double ResolveContainerElementWidth(HtmlComputedStyle style, double containingWidth, double fontSize, double rootFontSize, MediaEnvironment environment) {
        string width = style.GetValue("width");
        ResolveContainerInsets(style, containingWidth, fontSize, rootFontSize, environment, out double horizontalInsets, out _);
        if (HtmlRenderCssValues.TryLength(width, containingWidth, fontSize, rootFontSize, environment.Width, environment.Height, out double resolved)
            && resolved >= 0D) {
            return Math.Max(0D, resolved - (IsBorderBox(style) ? horizontalInsets : 0D));
        }
        return Math.Max(0D, containingWidth - horizontalInsets);
    }

    private static double? ResolveContainerElementHeight(HtmlComputedStyle style, double containingWidth, double? containingHeight, double fontSize, double rootFontSize, MediaEnvironment environment) {
        string height = style.GetValue("height");
        if (!HtmlRenderCssValues.TryLength(height, containingHeight ?? double.NaN, fontSize, rootFontSize, environment.Width, environment.Height, out double resolved)
            || resolved < 0D) return null;
        ResolveContainerInsets(style, containingWidth, fontSize, rootFontSize, environment, out _, out double verticalInsets);
        return Math.Max(0D, resolved - (IsBorderBox(style) ? verticalInsets : 0D));
    }

    private static bool IsBorderBox(HtmlComputedStyle style) =>
        string.Equals(style.GetValue("box-sizing").Trim(), "border-box", StringComparison.OrdinalIgnoreCase);

    private static void ResolveContainerInsets(
        HtmlComputedStyle style,
        double containingWidth,
        double fontSize,
        double rootFontSize,
        MediaEnvironment environment,
        out double horizontal,
        out double vertical) {
        double top = 0D;
        double right = 0D;
        double bottom = 0D;
        double left = 0D;
        string padding = style.GetValue("padding");
        if (padding.Length > 0) {
            HtmlRenderCssValues.ApplyBoxShorthand(
                padding,
                containingWidth,
                fontSize,
                rootFontSize,
                environment.Width,
                environment.Height,
                ref top,
                ref right,
                ref bottom,
                ref left);
        }
        ApplyContainerInset(style.GetValue("padding-top"), containingWidth, fontSize, rootFontSize, environment, ref top);
        ApplyContainerInset(style.GetValue("padding-right"), containingWidth, fontSize, rootFontSize, environment, ref right);
        ApplyContainerInset(style.GetValue("padding-bottom"), containingWidth, fontSize, rootFontSize, environment, ref bottom);
        ApplyContainerInset(style.GetValue("padding-left"), containingWidth, fontSize, rootFontSize, environment, ref left);

        double borderHorizontal = 0D;
        double borderVertical = 0D;
        if (HtmlCssBoxStrokeParser.TryParseBorder(
                style,
                containingWidth,
                fontSize,
                rootFontSize,
                environment.Width,
                environment.Height,
                OfficeIMO.Drawing.OfficeColor.Black,
                out HtmlRenderBorderEdges borders,
                out _)) {
            borderHorizontal = borders.Left.LayoutWidth + borders.Right.LayoutWidth;
            borderVertical = borders.Top.LayoutWidth + borders.Bottom.LayoutWidth;
        }
        horizontal = Math.Max(0D, left) + Math.Max(0D, right) + borderHorizontal;
        vertical = Math.Max(0D, top) + Math.Max(0D, bottom) + borderVertical;
    }

    private static void ApplyContainerInset(
        string value,
        double reference,
        double fontSize,
        double rootFontSize,
        MediaEnvironment environment,
        ref double target) {
        if (value.Length > 0
            && HtmlRenderCssValues.TryLength(value, reference, fontSize, rootFontSize, environment.Width, environment.Height, out double parsed)) {
            target = Math.Max(0D, parsed);
        }
    }

    private static double ResolveContainerFontSize(
        HtmlComputedStyle style,
        MediaEnvironment environment,
        double inheritedFontSize = 16D,
        double rootFontSize = 16D) =>
        HtmlRenderCssValues.TryLength(style.GetValue("font-size"), inheritedFontSize, inheritedFontSize, rootFontSize, environment.Width, environment.Height, out double fontSize)
            && fontSize > 0D
            ? fontSize
            : inheritedFontSize;

    private static bool EvaluateContainerCondition(string condition, ContainerQueryContext context, MediaEnvironment environment) {
        string normalized = condition.Trim();
        if (normalized.Length == 0) return false;
        if (StartsWithLogicalNot(normalized)) return !EvaluateContainerCondition(normalized.Substring(3).Trim(), context, environment);

        IReadOnlyList<string> orParts = SplitTopLevelLogical(normalized, "or").ToList();
        if (orParts.Count > 1) return orParts.Any(part => EvaluateContainerCondition(part, context, environment));
        IReadOnlyList<string> andParts = SplitTopLevelLogical(normalized, "and").ToList();
        if (andParts.Count > 1) return andParts.All(part => EvaluateContainerCondition(part, context, environment));

        if (normalized[0] == '(' && FindMatchingParenthesis(normalized, 0) == normalized.Length - 1) {
            normalized = normalized.Substring(1, normalized.Length - 2).Trim();
        }
        if (normalized.StartsWith("style(", StringComparison.OrdinalIgnoreCase)
            && normalized.EndsWith(")", StringComparison.Ordinal)) {
            return EvaluateContainerStyleQuery(normalized.Substring(6, normalized.Length - 7), context, environment);
        }
        return EvaluateContainerSizeFeature(normalized, context, environment);
    }

    private static bool EvaluateContainerStyleQuery(string query, ContainerQueryContext context, MediaEnvironment environment) {
        int colon = query.IndexOf(':');
        string name = (colon < 0 ? query : query.Substring(0, colon)).Trim();
        if (name.Length == 0 || !context.Properties.TryGetValue(name, out string? actual)) return false;
        if (colon < 0) return actual.Length > 0;
        string expected = query.Substring(colon + 1).Trim();
        if (name.StartsWith("--", StringComparison.Ordinal)) {
            string? LookupCustomProperty(string property) => context.Properties.TryGetValue(property, out string? value) ? value : null;
            if (!HtmlCssCustomPropertyResolver.TryResolve(actual, LookupCustomProperty, out string resolvedActual)
                || !HtmlCssCustomPropertyResolver.TryResolve(expected, LookupCustomProperty, out string resolvedExpected)) {
                return false;
            }
            return string.Equals(
                string.Join(" ", HtmlRenderCssValues.SplitWhitespace(resolvedActual)),
                string.Join(" ", HtmlRenderCssValues.SplitWhitespace(resolvedExpected)),
                StringComparison.Ordinal);
        }
        if (HtmlRenderCssValues.TryColor(actual, out OfficeIMO.Drawing.OfficeColor actualColor)
            && HtmlRenderCssValues.TryColor(expected, out OfficeIMO.Drawing.OfficeColor expectedColor)) {
            return actualColor == expectedColor;
        }
        if (HtmlRenderCssValues.TryLength(actual, context.Width, context.FontSize, context.RootFontSize, environment.Width, environment.Height, out double actualLength)
            && HtmlRenderCssValues.TryLength(expected, context.Width, context.FontSize, context.RootFontSize, environment.Width, environment.Height, out double expectedLength)) {
            return Math.Abs(actualLength - expectedLength) <= 0.000001D;
        }
        return string.Equals(
            string.Join(" ", HtmlRenderCssValues.SplitWhitespace(actual)),
            string.Join(" ", HtmlRenderCssValues.SplitWhitespace(expected)),
            StringComparison.OrdinalIgnoreCase);
    }

    private static bool EvaluateContainerSizeFeature(string feature, ContainerQueryContext context, MediaEnvironment environment) {
        int colon = feature.IndexOf(':');
        if (colon >= 0) {
            string name = feature.Substring(0, colon).Trim().ToLowerInvariant();
            string expectedText = feature.Substring(colon + 1).Trim();
            if (name == "orientation") {
                if (!context.Height.HasValue) return false;
                string actualOrientation = context.Width >= context.Height.Value ? "landscape" : "portrait";
                return string.Equals(actualOrientation, expectedText, StringComparison.OrdinalIgnoreCase);
            }
            bool minimum = name.StartsWith("min-", StringComparison.Ordinal);
            bool maximum = name.StartsWith("max-", StringComparison.Ordinal);
            string baseName = minimum || maximum ? name.Substring(4) : name;
            if (!TryGetContainerFeatureValue(baseName, context, out double actual)
                || !TryParseContainerFeatureValue(baseName, expectedText, context, environment, out double expected)) return false;
            return minimum ? actual >= expected : maximum ? actual <= expected : Math.Abs(actual - expected) <= 0.000001D;
        }

        IReadOnlyList<string> parts = HtmlRenderCssValues.SplitWhitespace(feature);
        if (parts.Count == 3) {
            if (TryGetContainerFeatureValue(parts[0], context, out double actual)
                && TryParseContainerFeatureValue(parts[0], parts[2], context, environment, out double expected)) {
                return CompareContainerValues(actual, expected, parts[1]);
            }
            if (TryGetContainerFeatureValue(parts[2], context, out actual)
                && TryParseContainerFeatureValue(parts[2], parts[0], context, environment, out expected)) {
                return CompareContainerValues(expected, actual, parts[1]);
            }
        }
        if (parts.Count == 5
            && TryGetContainerFeatureValue(parts[2], context, out double middle)
            && TryParseContainerFeatureValue(parts[2], parts[0], context, environment, out double lower)
            && TryParseContainerFeatureValue(parts[2], parts[4], context, environment, out double upper)) {
            return CompareContainerValues(lower, middle, parts[1]) && CompareContainerValues(middle, upper, parts[3]);
        }
        return false;
    }

    private static bool TryGetContainerFeatureValue(string name, ContainerQueryContext context, out double value) {
        string normalized = name.Trim().ToLowerInvariant();
        if (normalized == "width" || normalized == "inline-size") {
            value = context.Width;
            return true;
        }
        if (normalized == "height" || normalized == "block-size") {
            value = context.Height ?? 0D;
            return context.Height.HasValue;
        }
        if (normalized == "aspect-ratio" && context.Height.HasValue && context.Height.Value > 0D) {
            value = context.Width / context.Height.Value;
            return true;
        }
        value = 0D;
        return false;
    }

    private static bool TryParseContainerFeatureValue(
        string feature,
        string text,
        ContainerQueryContext context,
        MediaEnvironment environment,
        out double value) {
        if (string.Equals(feature.Trim(), "aspect-ratio", StringComparison.OrdinalIgnoreCase)) {
            IReadOnlyList<string> ratio = HtmlRenderCssValues.SplitTopLevel(text, '/');
            if (ratio.Count == 2
                && double.TryParse(ratio[0], NumberStyles.Float, CultureInfo.InvariantCulture, out double numerator)
                && double.TryParse(ratio[1], NumberStyles.Float, CultureInfo.InvariantCulture, out double denominator)
                && numerator >= 0D && denominator > 0D) {
                value = numerator / denominator;
                return true;
            }
        }
        return HtmlRenderCssValues.TryLength(
            text,
            context.Width,
            context.FontSize,
            context.RootFontSize,
            environment.Width,
            environment.Height,
            context.Width,
            context.Height ?? double.NaN,
            out value);
    }

    private static bool CompareContainerValues(double left, double right, string operation) {
        switch (operation) {
            case "<": return left < right;
            case "<=": return left <= right;
            case ">": return left > right;
            case ">=": return left >= right;
            case "=": return Math.Abs(left - right) <= 0.000001D;
            default: return false;
        }
    }
}
