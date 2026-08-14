namespace OfficeIMO.Html;

public static partial class HtmlComputedStyleEngine {
    private static bool IsSupportsRule(AngleSharp.Css.Dom.ICssRule rule) =>
        rule is AngleSharp.Css.Dom.ICssSupportsRule;

    private static string GetConditionText(AngleSharp.Css.Dom.ICssRule rule) =>
        (rule as AngleSharp.Css.Dom.ICssSupportsRule)?.ConditionText ?? string.Empty;

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
        if (string.Equals(propertyName, "caption-side", StringComparison.OrdinalIgnoreCase)) {
            return IsKnownKeyword(normalized, "top", "bottom");
        }
        if (string.Equals(propertyName, "table-layout", StringComparison.OrdinalIgnoreCase)) {
            return IsKnownKeyword(normalized, "auto", "fixed");
        }
        if (string.Equals(propertyName, "border-collapse", StringComparison.OrdinalIgnoreCase)) {
            return IsKnownKeyword(normalized, "separate", "collapse");
        }
        if (string.Equals(propertyName, "border-spacing", StringComparison.OrdinalIgnoreCase)) {
            return HtmlCssTableParser.TryParseBorderSpacing(normalized, 16D, 16D, 100D, 100D, out _, out _);
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
        if (string.Equals(propertyName, "overflow-clip-margin", StringComparison.OrdinalIgnoreCase)) {
            return HtmlCssOverflowClipMarginParser.TryParse(normalized, 16D, 16D, 100D, 100D, out _, out _);
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
        if (string.Equals(propertyName, "object-fit", StringComparison.OrdinalIgnoreCase)) {
            return HtmlCssReplacedElementParser.IsSupportedObjectFitSyntax(normalized);
        }
        if (string.Equals(propertyName, "image-orientation", StringComparison.OrdinalIgnoreCase)) {
            return HtmlCssReplacedElementParser.IsSupportedImageOrientationSyntax(normalized);
        }
        if (string.Equals(propertyName, "image-resolution", StringComparison.OrdinalIgnoreCase)) {
            return HtmlCssReplacedElementParser.IsSupportedImageResolutionSyntax(normalized);
        }
        if (string.Equals(propertyName, "object-position", StringComparison.OrdinalIgnoreCase)) {
            return HtmlCssReplacedElementParser.IsSupportedObjectPositionSyntax(normalized);
        }
        if (string.Equals(propertyName, "aspect-ratio", StringComparison.OrdinalIgnoreCase)) {
            return HtmlCssReplacedElementParser.IsSupportedAspectRatioSyntax(normalized);
        }
        if (string.Equals(propertyName, "transform", StringComparison.OrdinalIgnoreCase)) {
            return HtmlCssTransformParser.IsSupportedTransformSyntax(normalized);
        }
        if (string.Equals(propertyName, "transform-origin", StringComparison.OrdinalIgnoreCase)) {
            return HtmlCssTransformParser.IsSupportedOriginSyntax(normalized);
        }
        if (string.Equals(propertyName, "clip-path", StringComparison.OrdinalIgnoreCase)) {
            return HtmlCssClipPathParser.IsSupportedSyntax(normalized);
        }
        if (string.Equals(propertyName, "-officeimo-pdf-tag-type", StringComparison.OrdinalIgnoreCase)) {
            return HtmlRenderStyleResolver.IsSupportedPdfTagType(normalized);
        }
        if (string.Equals(propertyName, "bookmark-level", StringComparison.OrdinalIgnoreCase)) {
            return normalized == "none"
                || int.TryParse(normalized, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int level)
                && level >= 1 && level <= 64;
        }
        if (string.Equals(propertyName, "bookmark-state", StringComparison.OrdinalIgnoreCase)) {
            return normalized == "open" || normalized == "closed";
        }
        if (string.Equals(propertyName, "bookmark-label", StringComparison.OrdinalIgnoreCase)) {
            return normalized == "content(text)" || IsQuotedCssString(value.Trim());
        }
        if (string.Equals(propertyName, "border-radius", StringComparison.OrdinalIgnoreCase)) {
            return HtmlCssBorderRadiusParser.IsSupportedShorthandSyntax(normalized);
        }
        if (string.Equals(propertyName, "box-shadow", StringComparison.OrdinalIgnoreCase)) {
            return HtmlCssBoxShadowParser.IsSupportedSyntax(normalized);
        }
        if (string.Equals(propertyName, "border", StringComparison.OrdinalIgnoreCase)) {
            return HtmlCssBoxStrokeParser.IsSupportedBorderSyntax(normalized);
        }
        if (string.Equals(propertyName, "border-top", StringComparison.OrdinalIgnoreCase)
            || string.Equals(propertyName, "border-right", StringComparison.OrdinalIgnoreCase)
            || string.Equals(propertyName, "border-bottom", StringComparison.OrdinalIgnoreCase)
            || string.Equals(propertyName, "border-left", StringComparison.OrdinalIgnoreCase)) {
            return HtmlCssBoxStrokeParser.IsSupportedBorderSyntax(normalized);
        }
        if (string.Equals(propertyName, "border-top-width", StringComparison.OrdinalIgnoreCase)
            || string.Equals(propertyName, "border-right-width", StringComparison.OrdinalIgnoreCase)
            || string.Equals(propertyName, "border-bottom-width", StringComparison.OrdinalIgnoreCase)
            || string.Equals(propertyName, "border-left-width", StringComparison.OrdinalIgnoreCase)) {
            return HtmlCssBoxStrokeParser.IsSupportedSideWidthSyntax(normalized);
        }
        if (string.Equals(propertyName, "border-top-style", StringComparison.OrdinalIgnoreCase)
            || string.Equals(propertyName, "border-right-style", StringComparison.OrdinalIgnoreCase)
            || string.Equals(propertyName, "border-bottom-style", StringComparison.OrdinalIgnoreCase)
            || string.Equals(propertyName, "border-left-style", StringComparison.OrdinalIgnoreCase)) {
            return HtmlCssBoxStrokeParser.IsSupportedSideStyleSyntax(normalized);
        }
        if (string.Equals(propertyName, "border-top-color", StringComparison.OrdinalIgnoreCase)
            || string.Equals(propertyName, "border-right-color", StringComparison.OrdinalIgnoreCase)
            || string.Equals(propertyName, "border-bottom-color", StringComparison.OrdinalIgnoreCase)
            || string.Equals(propertyName, "border-left-color", StringComparison.OrdinalIgnoreCase)) {
            return HtmlCssBoxStrokeParser.IsSupportedSideColorSyntax(normalized);
        }
        if (string.Equals(propertyName, "border-width", StringComparison.OrdinalIgnoreCase)) {
            return HtmlCssBoxStrokeParser.IsSupportedWidthSyntax(normalized);
        }
        if (string.Equals(propertyName, "outline-width", StringComparison.OrdinalIgnoreCase)) {
            return HtmlCssBoxStrokeParser.IsSupportedSideWidthSyntax(normalized);
        }
        if (string.Equals(propertyName, "border-style", StringComparison.OrdinalIgnoreCase)) {
            return HtmlCssBoxStrokeParser.IsSupportedStyleSyntax(normalized);
        }
        if (string.Equals(propertyName, "outline-style", StringComparison.OrdinalIgnoreCase)) {
            return HtmlCssBoxStrokeParser.IsSupportedSideStyleSyntax(normalized);
        }
        if (string.Equals(propertyName, "border-color", StringComparison.OrdinalIgnoreCase)) {
            return HtmlCssBoxStrokeParser.IsSupportedColorSyntax(normalized);
        }
        if (string.Equals(propertyName, "outline-color", StringComparison.OrdinalIgnoreCase)) {
            return HtmlCssBoxStrokeParser.IsSupportedSideColorSyntax(normalized);
        }
        if (string.Equals(propertyName, "outline", StringComparison.OrdinalIgnoreCase)) {
            return HtmlCssBoxStrokeParser.IsSupportedOutlineSyntax(normalized);
        }
        if (string.Equals(propertyName, "outline-offset", StringComparison.OrdinalIgnoreCase)) {
            return !normalized.EndsWith("%", StringComparison.Ordinal)
                && TryValidateCssLength(normalized, out _);
        }
        if (string.Equals(propertyName, "border-top-left-radius", StringComparison.OrdinalIgnoreCase)
            || string.Equals(propertyName, "border-top-right-radius", StringComparison.OrdinalIgnoreCase)
            || string.Equals(propertyName, "border-bottom-right-radius", StringComparison.OrdinalIgnoreCase)
            || string.Equals(propertyName, "border-bottom-left-radius", StringComparison.OrdinalIgnoreCase)) {
            return HtmlCssBorderRadiusParser.IsSupportedCornerSyntax(normalized);
        }
        return IsSupportedDeclarationValue(propertyName, value);
    }

    private static bool IsPositiveCssLength(string value) {
        return HtmlRenderCssValues.HasExplicitLengthSyntax(value, allowPercentage: false, allowUnitlessZero: false)
            && TryValidateCssLength(value, out double length)
            && length > 0D;
    }

    private static bool IsNonNegativeCssLength(string value) {
        return HtmlRenderCssValues.HasExplicitLengthSyntax(value, allowPercentage: false, allowUnitlessZero: true)
            && TryValidateCssLength(value, out double length)
            && length >= 0D;
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

        string rawNormalized = value.Trim().ToLowerInvariant();
        if (IsCssWideKeyword(rawNormalized)) {
            return true;
        }
        string normalized = rawNormalized;
        switch (propertyName.ToLowerInvariant()) {
            case "position":
                return IsKnownKeyword(normalized, "static", "relative", "absolute", "fixed", "sticky")
                    || HtmlCssRunningElementParser.TryParsePosition(value, out _);
            case "container-type":
                return IsKnownKeyword(normalized, "normal", "size", "inline-size");
            case "container-name":
                return normalized == "none" || TryParseContainerNameList(normalized, out _);
            case "container":
                int slash = normalized.IndexOf('/');
                if (slash < 0) return normalized == "none" || TryParseContainerNameList(normalized, out _);
                string containerNames = normalized.Substring(0, slash).Trim();
                string containerType = normalized.Substring(slash + 1).Trim();
                return containerNames.Length > 0
                    && (containerNames == "none" || TryParseContainerNameList(containerNames, out _))
                    && IsKnownKeyword(containerType, "normal", "size", "inline-size");
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
            case "hyphens":
                return IsKnownKeyword(normalized, "none", "manual", "auto");
            case "hyphenate-character":
                return IsSupportedHyphenateCharacterSyntax(value.Trim());
            case "hyphenate-limit-chars":
                return IsSupportedHyphenateLimitCharsSyntax(normalized);
            case "hyphenate-limit-lines":
                return normalized == "no-limit"
                    || int.TryParse(normalized, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int hyphenatedLines)
                    && hyphenatedLines > 0;
            case "hyphenate-limit-last":
                return IsKnownKeyword(normalized, "none", "always");
            case "hyphenate-limit-zone":
                return IsNonNegativeCssLengthOrPercentage(normalized);
            case "text-overflow":
                return IsKnownKeyword(normalized, "clip", "ellipsis");
            case "line-clamp":
            case "-webkit-line-clamp":
                return normalized == "none"
                    || int.TryParse(normalized, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int clampedLines)
                    && clampedLines > 0;
            case "tab-size":
                return double.TryParse(normalized, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double tabCount)
                    ? tabCount >= 0D && !double.IsNaN(tabCount) && !double.IsInfinity(tabCount)
                    : IsNonNegativeCssLength(normalized);
            case "letter-spacing":
            case "word-spacing":
                return normalized == "normal"
                    || HtmlRenderCssValues.HasExplicitLengthSyntax(normalized, allowPercentage: false, allowUnitlessZero: true)
                    && TryValidateCssLength(normalized, out _);
            case "image-orientation":
                return HtmlCssReplacedElementParser.IsSupportedImageOrientationSyntax(normalized);
            case "image-resolution":
                return HtmlCssReplacedElementParser.IsSupportedImageResolutionSyntax(normalized);
            default:
                return !normalized.StartsWith("not-a-real", StringComparison.Ordinal);
        }
    }

    private static bool IsCssWideKeyword(string value) =>
        IsKnownKeyword(value, "inherit", "initial", "revert", "revert-layer", "unset");

    private static bool IsQuotedCssString(string value) =>
        value.Length >= 2 && (value[0] == '\'' || value[0] == '"') && value[value.Length - 1] == value[0];

    private static bool IsSupportedHyphenateCharacterSyntax(string value) {
        if (string.Equals(value, "auto", StringComparison.OrdinalIgnoreCase)) return true;
        if (value.Length < 2 || value[0] != value[value.Length - 1] || value[0] != '\'' && value[0] != '"') return false;
        return HtmlCssEscapeDecoder.Decode(value.Substring(1, value.Length - 2)).Length <= 8;
    }

    private static bool IsSupportedHyphenateLimitCharsSyntax(string value) {
        IReadOnlyList<string> parts = HtmlRenderCssValues.SplitWhitespace(value);
        return parts.Count is >= 1 and <= 3
            && parts.All(part => part == "auto"
                || int.TryParse(part, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int parsed) && parsed > 0);
    }

    private static bool IsNonNegativeCssLengthOrPercentage(string value) =>
        HtmlRenderCssValues.HasExplicitLengthSyntax(value, allowPercentage: true, allowUnitlessZero: true)
        && TryValidateCssLength(value, out double length)
        && length >= 0D;

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
        IReadOnlyDictionary<string, string>? parentProperties,
        out IReadOnlyCollection<string> inheritedProperties,
        out IReadOnlyCollection<string> resetProperties) {
        var raw = new Dictionary<string, string>(HtmlCssPropertyNameComparer.Instance);
        var inherited = new HashSet<string>(HtmlCssPropertyNameComparer.Instance);
        var reset = new HashSet<string>(HtmlCssPropertyNameComparer.Instance);
        foreach (KeyValuePair<string, CascadedProperty> pair in properties) {
            CascadedProperty? effective = ResolveLayerRevert(pair.Value);
            if (effective?.HasValue == true) {
                raw[pair.Key] = effective.Value;
                if (ReferenceEquals(effective.Specificity, Specificity.Inherited) || effective.InheritsComputedValue) inherited.Add(pair.Key);
            }
            else if (effective?.RevertsLayer == true
                && IsInheritedProperty(pair.Key)
                && parentProperties != null
                && parentProperties.TryGetValue(pair.Key, out string? inheritedValue)) {
                raw[pair.Key] = inheritedValue;
                inherited.Add(pair.Key);
            } else {
                reset.Add(pair.Key);
            }
        }
        var resolved = new Dictionary<string, string>(HtmlCssPropertyNameComparer.Instance);
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

        inherited.IntersectWith(resolved.Keys);
        inheritedProperties = inherited;
        reset.ExceptWith(resolved.Keys);
        resetProperties = reset;
        return resolved;
    }

    private static bool TryValidateCssLength(string value, out double length) =>
        HtmlRenderCssValues.TryLength(value, 100D, 16D, 16D, 100D, 100D, out length);

    private static bool TryParseContainerNameList(string value, out IReadOnlyList<string> names) {
        var parsed = new List<string>();
        int cursor = 0;
        while (cursor < value.Length) {
            while (cursor < value.Length && char.IsWhiteSpace(value[cursor])) cursor++;
            if (cursor >= value.Length) break;
            if (!HtmlCssIdentifierParser.TryRead(value, ref cursor, out string identifier) || !IsContainerNameIdentifier(identifier)) {
                names = Array.Empty<string>();
                return false;
            }
            if (cursor < value.Length && !char.IsWhiteSpace(value[cursor])) {
                names = Array.Empty<string>();
                return false;
            }
            parsed.Add(identifier);
        }
        names = parsed.AsReadOnly();
        return parsed.Count > 0;
    }

    private static bool IsContainerNameIdentifier(string identifier) {
        return !IsKnownKeyword(identifier.ToLowerInvariant(), "none", "and", "or", "not", "default", "inherit", "initial", "revert", "revert-layer", "unset");
    }

    private static CascadedProperty? ResolveLayerRevert(CascadedProperty property) {
        if (!property.RevertsLayer) return property;

        var candidates = new List<CascadedProperty>(property.Alternatives.Count + 1) { property };
        candidates.AddRange(property.Alternatives);
        var revertedLayers = new HashSet<CascadeLayerOrder?>();
        while (true) {
            CascadedProperty? current = null;
            foreach (CascadedProperty candidate in candidates) {
                if (candidate.Specificity != Specificity.Inherited && revertedLayers.Contains(candidate.LayerOrder)) continue;
                if (current == null || ShouldReplace(current, candidate.IsImportant, candidate.Specificity, candidate.Order, candidate.LayerOrder)) {
                    current = candidate;
                }
            }
            if (current?.RevertsLayer != true) return current;
            if (current.LayerOrder == null) return current;
            revertedLayers.Add(current.LayerOrder);
        }
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
