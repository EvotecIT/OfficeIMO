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
        double inheritedFontSize,
        double rootFontSize,
        IReadOnlyList<ContainerQueryContext> contexts) {
        string type = ResolveContainerType(style);
        IReadOnlyList<string> names = ResolveContainerNames(style);
        var expanded = new List<ContainerQueryContext>(contexts.Count + 1);
        expanded.AddRange(contexts);
        expanded.Add(new ContainerQueryContext(names, type, width, height, fontSize, inheritedFontSize, rootFontSize, style.Properties));
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
        return TryParseContainerNameList(value, out IReadOnlyList<string> names) ? names : Array.Empty<string>();
    }

    private static void ResolveContainerUnitDimensions(
        IReadOnlyList<ContainerQueryContext> contexts,
        out double width,
        out double height) {
        width = double.NaN;
        height = double.NaN;
        for (int index = contexts.Count - 1; index >= 0; index--) {
            ContainerQueryContext context = contexts[index];
            if (double.IsNaN(width) && (context.Type == "inline-size" || context.Type == "size")) width = context.Width;
            if (double.IsNaN(height) && context.Type == "size" && context.Height.HasValue) height = context.Height.Value;
            if (!double.IsNaN(width) && !double.IsNaN(height)) break;
        }
    }

    private static double ResolveContainerElementWidth(
        HtmlComputedStyle style,
        double containingWidth,
        double fontSize,
        double rootFontSize,
        MediaEnvironment environment,
        double containerUnitWidth,
        double containerUnitHeight) {
        string width = style.GetValue("width");
        ResolveContainerInsets(style, containingWidth, fontSize, rootFontSize, environment, containerUnitWidth, containerUnitHeight, out double horizontalInsets, out _);
        bool borderBox = IsBorderBox(style);
        double contentWidth;
        if (HtmlRenderCssValues.TryLength(width, containingWidth, fontSize, rootFontSize, environment.Width, environment.Height, containerUnitWidth, containerUnitHeight, out double resolved)
            && resolved >= 0D) {
            contentWidth = Math.Max(0D, resolved - (borderBox ? horizontalInsets : 0D));
        } else {
            contentWidth = Math.Max(0D, containingWidth - horizontalInsets);
        }

        if (TryResolveContainerDimensionConstraint(style.GetValue("max-width"), containingWidth, fontSize, rootFontSize, environment, containerUnitWidth, containerUnitHeight, horizontalInsets, borderBox, out double maximum)) {
            contentWidth = Math.Min(contentWidth, maximum);
        }
        if (TryResolveContainerDimensionConstraint(style.GetValue("min-width"), containingWidth, fontSize, rootFontSize, environment, containerUnitWidth, containerUnitHeight, horizontalInsets, borderBox, out double minimum)) {
            contentWidth = Math.Max(contentWidth, minimum);
        }
        return contentWidth;
    }

    private static bool TryResolveContainerDimensionConstraint(
        string value,
        double containingSize,
        double fontSize,
        double rootFontSize,
        MediaEnvironment environment,
        double containerUnitWidth,
        double containerUnitHeight,
        double insets,
        bool borderBox,
        out double contentSize) {
        contentSize = 0D;
        if (!HtmlRenderCssValues.TryLength(value, containingSize, fontSize, rootFontSize, environment.Width, environment.Height, containerUnitWidth, containerUnitHeight, out double resolved)
            || resolved < 0D) {
            return false;
        }
        contentSize = Math.Max(0D, resolved - (borderBox ? insets : 0D));
        return true;
    }

    private static double? ResolveContainerElementHeight(
        HtmlComputedStyle style,
        double contentWidth,
        double containingWidth,
        double? containingHeight,
        double fontSize,
        double rootFontSize,
        MediaEnvironment environment,
        double containerUnitWidth,
        double containerUnitHeight) {
        string height = style.GetValue("height");
        ResolveContainerInsets(style, containingWidth, fontSize, rootFontSize, environment, containerUnitWidth, containerUnitHeight, out double horizontalInsets, out double verticalInsets);
        bool borderBox = IsBorderBox(style);
        double contentHeight = 0D;
        bool hasDefiniteHeight = false;
        if (HtmlRenderCssValues.TryLength(height, containingHeight ?? double.NaN, fontSize, rootFontSize, environment.Width, environment.Height, containerUnitWidth, containerUnitHeight, out double resolved)
            && resolved >= 0D) {
            contentHeight = Math.Max(0D, resolved - (borderBox ? verticalInsets : 0D));
            hasDefiniteHeight = true;
        } else if (HtmlCssReplacedElementParser.TryParseAspectRatio(style.GetValue("aspect-ratio"), out double? ratio, out _, out _)
            && ratio.HasValue) {
            double ratioWidth = borderBox ? contentWidth + horizontalInsets : contentWidth;
            double ratioHeight = ratioWidth / ratio.Value;
            contentHeight = Math.Max(0D, ratioHeight - (borderBox ? verticalInsets : 0D));
            hasDefiniteHeight = true;
        }
        double containingSize = containingHeight ?? double.NaN;
        if (hasDefiniteHeight
            && TryResolveContainerDimensionConstraint(style.GetValue("max-height"), containingSize, fontSize, rootFontSize, environment, containerUnitWidth, containerUnitHeight, verticalInsets, borderBox, out double maximum)) {
            contentHeight = Math.Min(contentHeight, maximum);
        }
        if (TryResolveContainerDimensionConstraint(style.GetValue("min-height"), containingSize, fontSize, rootFontSize, environment, containerUnitWidth, containerUnitHeight, verticalInsets, borderBox, out double minimum)) {
            contentHeight = Math.Max(contentHeight, minimum);
            hasDefiniteHeight = true;
        }
        return hasDefiniteHeight ? contentHeight : null;
    }

    private static bool IsBorderBox(HtmlComputedStyle style) =>
        string.Equals(style.GetValue("box-sizing").Trim(), "border-box", StringComparison.OrdinalIgnoreCase);

    private static void ResolveContainerInsets(
        HtmlComputedStyle style,
        double containingWidth,
        double fontSize,
        double rootFontSize,
        MediaEnvironment environment,
        double containerUnitWidth,
        double containerUnitHeight,
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
                containerUnitWidth,
                containerUnitHeight,
                ref top,
                ref right,
                ref bottom,
                ref left);
        }
        ApplyContainerInset(style.GetValue("padding-top"), containingWidth, fontSize, rootFontSize, environment, containerUnitWidth, containerUnitHeight, ref top);
        ApplyContainerInset(style.GetValue("padding-right"), containingWidth, fontSize, rootFontSize, environment, containerUnitWidth, containerUnitHeight, ref right);
        ApplyContainerInset(style.GetValue("padding-bottom"), containingWidth, fontSize, rootFontSize, environment, containerUnitWidth, containerUnitHeight, ref bottom);
        ApplyContainerInset(style.GetValue("padding-left"), containingWidth, fontSize, rootFontSize, environment, containerUnitWidth, containerUnitHeight, ref left);

        double borderHorizontal = 0D;
        double borderVertical = 0D;
        if (HtmlCssBoxStrokeParser.TryParseBorder(
                style,
                containingWidth,
                fontSize,
                rootFontSize,
                environment.Width,
                environment.Height,
                containerUnitWidth,
                containerUnitHeight,
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
        double containerUnitWidth,
        double containerUnitHeight,
        ref double target) {
        if (value.Length > 0
            && HtmlRenderCssValues.TryLength(value, reference, fontSize, rootFontSize, environment.Width, environment.Height, containerUnitWidth, containerUnitHeight, out double parsed)) {
            target = Math.Max(0D, parsed);
        }
    }

    private static double ResolveContainerFontSize(
        HtmlComputedStyle style,
        MediaEnvironment environment,
        double inheritedFontSize = 16D,
        double rootFontSize = 16D,
        double containerUnitWidth = double.NaN,
        double containerUnitHeight = double.NaN) =>
        HtmlRenderCssValues.TryLength(style.GetValue("font-size"), inheritedFontSize, inheritedFontSize, rootFontSize, environment.Width, environment.Height, containerUnitWidth, containerUnitHeight, out double fontSize)
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
            return EvaluateContainerCondition(normalized.Substring(1, normalized.Length - 2), context, environment);
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
        double fontReference = string.Equals(name, "font-size", StringComparison.OrdinalIgnoreCase)
            ? context.InheritedFontSize
            : context.FontSize;
        double percentageReference = string.Equals(name, "font-size", StringComparison.OrdinalIgnoreCase)
            ? context.InheritedFontSize
            : context.Width;
        if (HtmlRenderCssValues.TryLength(actual, percentageReference, fontReference, context.RootFontSize, environment.Width, environment.Height, context.Width, context.Height ?? double.NaN, out double actualLength)
            && HtmlRenderCssValues.TryLength(expected, percentageReference, fontReference, context.RootFontSize, environment.Width, environment.Height, context.Width, context.Height ?? double.NaN, out double expectedLength)) {
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
                string actualOrientation = context.Width > context.Height.Value ? "landscape" : "portrait";
                return string.Equals(actualOrientation, expectedText, StringComparison.OrdinalIgnoreCase);
            }
            bool minimum = name.StartsWith("min-", StringComparison.Ordinal);
            bool maximum = name.StartsWith("max-", StringComparison.Ordinal);
            string baseName = minimum || maximum ? name.Substring(4) : name;
            if (!TryGetContainerFeatureValue(baseName, context, out double actual)
                || !TryParseContainerFeatureValue(baseName, expectedText, context, environment, out double expected)) return false;
            return minimum ? actual >= expected : maximum ? actual <= expected : Math.Abs(actual - expected) <= 0.000001D;
        }

        if (!TryTokenizeContainerRange(feature, out IReadOnlyList<string> parts)) return false;
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

    private static bool TryTokenizeContainerRange(string feature, out IReadOnlyList<string> parts) {
        var tokens = new List<string>();
        int segmentStart = 0;
        int depth = 0;
        char quote = '\0';
        for (int index = 0; index < feature.Length; index++) {
            char current = feature[index];
            if (quote != '\0') {
                if (current == quote && (index == 0 || feature[index - 1] != '\\')) quote = '\0';
                continue;
            }
            if (current == '\'' || current == '"') {
                quote = current;
                continue;
            }
            if (current == '(') {
                depth++;
                continue;
            }
            if (current == ')' && depth > 0) {
                depth--;
                continue;
            }
            if (depth != 0 || current != '<' && current != '>' && current != '=') continue;

            string operand = feature.Substring(segmentStart, index - segmentStart).Trim();
            if (operand.Length == 0) {
                parts = Array.Empty<string>();
                return false;
            }
            tokens.Add(operand);
            int operatorLength = current != '=' && index + 1 < feature.Length && feature[index + 1] == '=' ? 2 : 1;
            tokens.Add(feature.Substring(index, operatorLength));
            index += operatorLength - 1;
            segmentStart = index + 1;
        }

        string finalOperand = feature.Substring(segmentStart).Trim();
        if (tokens.Count == 0 || finalOperand.Length == 0) {
            parts = Array.Empty<string>();
            return false;
        }
        tokens.Add(finalOperand);
        if (tokens.Count != 3 && tokens.Count != 5) {
            parts = Array.Empty<string>();
            return false;
        }
        parts = tokens;
        return true;
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
            if (HtmlCssReplacedElementParser.TryParseAspectRatio(text, out double? ratio, out bool prefersIntrinsic, out _)
                && ratio.HasValue
                && !prefersIntrinsic) {
                value = ratio.Value;
                return true;
            }
            value = 0D;
            return false;
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
