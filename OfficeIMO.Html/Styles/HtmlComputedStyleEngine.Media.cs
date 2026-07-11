using OfficeIMO.Drawing;
using System.Globalization;

namespace OfficeIMO.Html;

public static partial class HtmlComputedStyleEngine {
    /// <summary>
    /// Evaluates whether a CSS media query list applies to the requested OfficeIMO media context.
    /// </summary>
    public static bool IsApplicableMedia(string mediaText, HtmlCssMediaContext mediaContext) =>
        IsApplicableMedia(mediaText, MediaEnvironment.CreateDefault(mediaContext));

    /// <summary>
    /// Evaluates whether a CSS media query list applies to a media context and surface size.
    /// </summary>
    public static bool IsApplicableMedia(string mediaText, HtmlCssMediaContext mediaContext, double surfaceWidth, double surfaceHeight) {
        if (surfaceWidth <= 0D || double.IsNaN(surfaceWidth) || double.IsInfinity(surfaceWidth)) {
            throw new ArgumentOutOfRangeException(nameof(surfaceWidth));
        }
        if (surfaceHeight <= 0D || double.IsNaN(surfaceHeight) || double.IsInfinity(surfaceHeight)) {
            throw new ArgumentOutOfRangeException(nameof(surfaceHeight));
        }

        return IsApplicableMedia(mediaText, new MediaEnvironment(mediaContext, surfaceWidth, surfaceHeight));
    }

    private static bool IsApplicableMedia(string mediaText, MediaEnvironment environment) {
        if (string.IsNullOrWhiteSpace(mediaText)) return true;

        mediaText = StripCssCommentsOutsideStrings(mediaText);
        string activeType = environment.Context == HtmlCssMediaContext.Print ? "print" : "screen";
        foreach (string query in SplitSelectorList(mediaText)) {
            string normalized = query.Trim();
            if (TryConsumeMediaModifier(normalized, "not", out string negatedQuery)) {
                if (!IsPositiveMediaQueryApplicable(negatedQuery, activeType, environment)) return true;
                continue;
            }

            if (IsPositiveMediaQueryApplicable(normalized, activeType, environment)) return true;
        }

        return false;
    }

    private static bool IsPositiveMediaQueryApplicable(string mediaQuery, string activeType, MediaEnvironment environment) {
        return AreMediaFeaturesApplicable(mediaQuery, environment)
            && (ContainsMediaType(mediaQuery, "all") || ContainsMediaType(mediaQuery, activeType) || !ContainsExplicitMediaType(mediaQuery));
    }

    private static bool ContainsMediaType(string mediaQuery, string mediaType) {
        foreach (string token in mediaQuery.Split(new[] { ' ', '\t', '\r', '\n', '\f' }, StringSplitOptions.RemoveEmptyEntries)) {
            if (string.Equals(token.Trim(), mediaType, StringComparison.OrdinalIgnoreCase)) return true;
        }

        return false;
    }

    private static bool ContainsExplicitMediaType(string mediaQuery) => TryReadExplicitMediaType(mediaQuery, out _);

    private static bool TryReadExplicitMediaType(string mediaQuery, out string mediaType) {
        mediaType = string.Empty;
        string normalized = mediaQuery.TrimStart();
        if (TryConsumeMediaModifier(normalized, "not", out string withoutNot)) {
            normalized = withoutNot;
        } else if (TryConsumeMediaModifier(normalized, "only", out string withoutOnly)) {
            normalized = withoutOnly;
        }

        if (normalized.Length == 0 || normalized[0] == '(') return false;

        int cursor = 0;
        while (cursor < normalized.Length && (char.IsLetterOrDigit(normalized[cursor]) || normalized[cursor] == '-' || normalized[cursor] == '_')) cursor++;
        if (cursor == 0) return false;

        string token = normalized.Substring(0, cursor);
        if (string.Equals(token, "and", StringComparison.OrdinalIgnoreCase)
            || string.Equals(token, "or", StringComparison.OrdinalIgnoreCase)) return false;

        mediaType = token;
        return true;
    }

    private static bool TryConsumeMediaModifier(string mediaQuery, string modifier, out string remaining) {
        remaining = mediaQuery;
        if (mediaQuery.Length <= modifier.Length || !mediaQuery.StartsWith(modifier, StringComparison.OrdinalIgnoreCase)) return false;
        if (!char.IsWhiteSpace(mediaQuery[modifier.Length])) return false;
        remaining = mediaQuery.Substring(modifier.Length + 1).TrimStart();
        return true;
    }

    private static bool HasMediaFeatureConstraint(string mediaQuery) {
        return mediaQuery.IndexOf("(", StringComparison.Ordinal) >= 0
            || mediaQuery.IndexOf(":", StringComparison.Ordinal) >= 0;
    }

    private static bool AreMediaFeaturesApplicable(string mediaQuery, MediaEnvironment environment) {
        int index = 0;
        bool foundFeature = false;
        while (index < mediaQuery.Length) {
            int open = mediaQuery.IndexOf('(', index);
            if (open < 0) break;
            int close = FindMatchingParenthesis(mediaQuery, open);
            if (close <= open) return false;

            foundFeature = true;
            string feature = mediaQuery.Substring(open + 1, close - open - 1).Trim().ToLowerInvariant();
            if (!IsMediaFeatureApplicable(feature, environment)) return false;
            index = close + 1;
        }

        return foundFeature || !HasMediaFeatureConstraint(mediaQuery);
    }

    private static bool IsMediaFeatureApplicable(string feature, MediaEnvironment environment) {
        if (feature.Length == 0 || feature.IndexOf("not-a-real", StringComparison.Ordinal) >= 0) return false;

        if (feature.StartsWith("color", StringComparison.Ordinal)
            || feature.StartsWith("min-color", StringComparison.Ordinal)
            || feature.StartsWith("monochrome", StringComparison.Ordinal)) return true;

        if (TryEvaluateMediaLengthFeature(feature, environment, out bool lengthApplies)) return lengthApplies;

        if (feature.StartsWith("orientation", StringComparison.Ordinal)) {
            int colon = feature.IndexOf(':');
            if (colon < 0) return false;
            string value = feature.Substring(colon + 1).Trim();
            bool landscape = environment.Width >= environment.Height;
            return string.Equals(value, landscape ? "landscape" : "portrait", StringComparison.Ordinal);
        }

        if (feature.StartsWith("resolution", StringComparison.Ordinal)) return true;

        if (environment.Context != HtmlCssMediaContext.Print
            && (feature.StartsWith("hover", StringComparison.Ordinal)
                || feature.StartsWith("pointer", StringComparison.Ordinal))) return true;

        return false;
    }

    private static bool TryEvaluateMediaLengthFeature(string feature, MediaEnvironment environment, out bool applies) {
        int colon = feature.IndexOf(':');
        if (colon < 0) {
            applies = false;
            return false;
        }

        string name = feature.Substring(0, colon).Trim();
        bool recognized = name == "width" || name == "height"
            || name == "min-width" || name == "min-height"
            || name == "max-width" || name == "max-height";
        if (!recognized) {
            applies = false;
            return false;
        }

        string value = feature.Substring(colon + 1).Trim();
        if (!TryParseMediaLength(value, environment, out double expected)) {
            applies = false;
            return true;
        }

        double actual = name.EndsWith("width", StringComparison.Ordinal) ? environment.Width : environment.Height;
        applies = name.StartsWith("min-", StringComparison.Ordinal)
            ? actual >= expected
            : name.StartsWith("max-", StringComparison.Ordinal)
                ? actual <= expected
                : Math.Abs(actual - expected) <= 0.000001D;
        return true;
    }

    private static bool TryParseMediaLength(string value, MediaEnvironment environment, out double result) {
        result = 0D;
        if (value.Length == 0) return false;

        int cursor = 0;
        if (value[cursor] == '+' || value[cursor] == '-') cursor++;
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
        if (!hasDigit) return false;

        string number = value.Substring(0, cursor);
        string unit = value.Substring(cursor).Trim().ToLowerInvariant();
        if (!double.TryParse(number, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed) || parsed < 0D) return false;
        if (unit.Length == 0) {
            if (Math.Abs(parsed) > double.Epsilon) return false;
            return true;
        }

        double multiplier;
        switch (unit) {
            case "px": multiplier = 1D; break;
            case "em":
            case "rem":
            case "pc": multiplier = 16D; break;
            case "vw": multiplier = environment.Width / 100D; break;
            case "vh": multiplier = environment.Height / 100D; break;
            case "vmin": multiplier = Math.Min(environment.Width, environment.Height) / 100D; break;
            case "vmax": multiplier = Math.Max(environment.Width, environment.Height) / 100D; break;
            case "in": multiplier = HtmlRenderOptions.CssPixelsPerInch; break;
            case "cm": multiplier = HtmlRenderOptions.CssPixelsPerInch / 2.54D; break;
            case "mm": multiplier = HtmlRenderOptions.CssPixelsPerInch / 25.4D; break;
            case "q": multiplier = HtmlRenderOptions.CssPixelsPerInch / 101.6D; break;
            case "pt": multiplier = HtmlRenderOptions.CssPixelsPerInch / 72D; break;
            default: return false;
        }

        result = parsed * multiplier;
        return !double.IsNaN(result) && !double.IsInfinity(result);
    }

    private readonly struct MediaEnvironment {
        internal MediaEnvironment(HtmlCssMediaContext context, double width, double height) {
            Context = context;
            Width = width;
            Height = height;
        }

        internal HtmlCssMediaContext Context { get; }
        internal double Width { get; }
        internal double Height { get; }

        internal static MediaEnvironment CreateDefault(HtmlCssMediaContext context) {
            if (context == HtmlCssMediaContext.Print) {
                return new MediaEnvironment(
                    context,
                    OfficePageSizes.A4.WidthInches * HtmlRenderOptions.CssPixelsPerInch,
                    OfficePageSizes.A4.HeightInches * HtmlRenderOptions.CssPixelsPerInch);
            }

            return new MediaEnvironment(context, 816D, 1056D);
        }
    }
}
