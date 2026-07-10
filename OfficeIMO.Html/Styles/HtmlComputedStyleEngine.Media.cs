using System.Globalization;

namespace OfficeIMO.Html;

public static partial class HtmlComputedStyleEngine {
    /// <summary>
    /// Evaluates whether a CSS media query list applies to the requested OfficeIMO media context.
    /// </summary>
    public static bool IsApplicableMedia(string mediaText, HtmlCssMediaContext mediaContext) {
        if (string.IsNullOrWhiteSpace(mediaText)) {
            return true;
        }

        mediaText = StripCssCommentsOutsideStrings(mediaText);
        string activeType = mediaContext == HtmlCssMediaContext.Print ? "print" : "screen";
        foreach (string query in SplitSelectorList(mediaText)) {
            string normalized = query.Trim();
            if (TryConsumeMediaModifier(normalized, "not", out string negatedQuery)) {
                if (!IsPositiveMediaQueryApplicable(negatedQuery, activeType, mediaContext)) {
                    return true;
                }

                continue;
            }

            if (IsPositiveMediaQueryApplicable(normalized, activeType, mediaContext)) {
                return true;
            }
        }

        return false;
    }

    private static bool IsPositiveMediaQueryApplicable(string mediaQuery, string activeType, HtmlCssMediaContext mediaContext) {
        return AreMediaFeaturesApplicable(mediaQuery, mediaContext)
            && (ContainsMediaType(mediaQuery, "all") || ContainsMediaType(mediaQuery, activeType) || !ContainsExplicitMediaType(mediaQuery));
    }

    private static bool ContainsMediaType(string mediaQuery, string mediaType) {
        foreach (string token in mediaQuery.Split(new[] { ' ', '\t', '\r', '\n', '\f' }, StringSplitOptions.RemoveEmptyEntries)) {
            if (string.Equals(token.Trim(), mediaType, StringComparison.OrdinalIgnoreCase)) {
                return true;
            }
        }

        return false;
    }

    private static bool ContainsExplicitMediaType(string mediaQuery) {
        return TryReadExplicitMediaType(mediaQuery, out _);
    }

    private static bool TryReadExplicitMediaType(string mediaQuery, out string mediaType) {
        mediaType = string.Empty;
        string normalized = mediaQuery.TrimStart();
        if (TryConsumeMediaModifier(normalized, "not", out string withoutNot)) {
            normalized = withoutNot;
        } else if (TryConsumeMediaModifier(normalized, "only", out string withoutOnly)) {
            normalized = withoutOnly;
        }

        if (normalized.Length == 0 || normalized[0] == '(') {
            return false;
        }

        int cursor = 0;
        while (cursor < normalized.Length && (char.IsLetterOrDigit(normalized[cursor]) || normalized[cursor] == '-' || normalized[cursor] == '_')) {
            cursor++;
        }

        if (cursor == 0) {
            return false;
        }

        string token = normalized.Substring(0, cursor);
        if (string.Equals(token, "and", StringComparison.OrdinalIgnoreCase)
            || string.Equals(token, "or", StringComparison.OrdinalIgnoreCase)) {
            return false;
        }

        mediaType = token;
        return true;
    }

    private static bool TryConsumeMediaModifier(string mediaQuery, string modifier, out string remaining) {
        remaining = mediaQuery;
        if (mediaQuery.Length <= modifier.Length || !mediaQuery.StartsWith(modifier, StringComparison.OrdinalIgnoreCase)) {
            return false;
        }

        char separator = mediaQuery[modifier.Length];
        if (!char.IsWhiteSpace(separator)) {
            return false;
        }

        remaining = mediaQuery.Substring(modifier.Length + 1).TrimStart();
        return true;
    }

    private static bool HasMediaFeatureConstraint(string mediaQuery) {
        return mediaQuery.IndexOf("(", StringComparison.Ordinal) >= 0
            || mediaQuery.IndexOf(":", StringComparison.Ordinal) >= 0;
    }

    private static bool AreMediaFeaturesApplicable(string mediaQuery, HtmlCssMediaContext mediaContext) {
        int index = 0;
        bool foundFeature = false;
        while (index < mediaQuery.Length) {
            int open = mediaQuery.IndexOf('(', index);
            if (open < 0) {
                break;
            }

            int close = FindMatchingParenthesis(mediaQuery, open);
            if (close <= open) {
                return false;
            }

            foundFeature = true;
            string feature = mediaQuery.Substring(open + 1, close - open - 1).Trim().ToLowerInvariant();
            if (!IsMediaFeatureApplicable(feature, mediaContext)) {
                return false;
            }

            index = close + 1;
        }

        return foundFeature || !HasMediaFeatureConstraint(mediaQuery);
    }

    private static bool IsMediaFeatureApplicable(string feature, HtmlCssMediaContext mediaContext) {
        if (feature.Length == 0 || feature.IndexOf("not-a-real", StringComparison.Ordinal) >= 0) {
            return false;
        }

        if (feature.StartsWith("color", StringComparison.Ordinal)
            || feature.StartsWith("min-color", StringComparison.Ordinal)
            || feature.StartsWith("monochrome", StringComparison.Ordinal)) {
            return true;
        }

        if (feature.StartsWith("max-width", StringComparison.Ordinal)
            || feature.StartsWith("max-height", StringComparison.Ordinal)) {
            return !IsZeroMaxMediaLength(feature);
        }

        if (feature.StartsWith("min-width", StringComparison.Ordinal)
            || feature.StartsWith("min-height", StringComparison.Ordinal)
            || feature.StartsWith("orientation", StringComparison.Ordinal)
            || feature.StartsWith("resolution", StringComparison.Ordinal)) {
            return true;
        }

        if (mediaContext != HtmlCssMediaContext.Print
            && (feature.StartsWith("hover", StringComparison.Ordinal)
                || feature.StartsWith("pointer", StringComparison.Ordinal))) {
            return true;
        }

        return false;
    }

    private static bool IsZeroMaxMediaLength(string feature) {
        int colon = feature.IndexOf(':');
        if (colon < 0) {
            return false;
        }

        string value = feature.Substring(colon + 1).Trim();
        if (value.Length == 0) {
            return false;
        }

        int cursor = 0;
        if (value[cursor] == '+' || value[cursor] == '-') {
            cursor++;
        }

        bool hasDigit = false;
        while (cursor < value.Length && char.IsDigit(value[cursor])) {
            hasDigit = true;
            cursor++;
        }

        if (cursor < value.Length && value[cursor] == '.') {
            cursor++;
            while (cursor < value.Length && char.IsDigit(value[cursor])) {
                hasDigit = true;
                cursor++;
            }
        }

        if (!hasDigit) {
            return false;
        }

        string number = value.Substring(0, cursor);
        string unit = value.Substring(cursor).Trim();
        if (!double.TryParse(number, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed)
            || Math.Abs(parsed) > double.Epsilon) {
            return false;
        }

        return unit.Length == 0
            || unit == "px"
            || unit == "em"
            || unit == "rem"
            || unit == "vw"
            || unit == "vh"
            || unit == "vmin"
            || unit == "vmax"
            || unit == "cm"
            || unit == "mm"
            || unit == "q"
            || unit == "in"
            || unit == "pc"
            || unit == "pt";
    }


}
