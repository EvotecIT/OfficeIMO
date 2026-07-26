using OfficeIMO.Drawing;
using System.Globalization;

namespace OfficeIMO.Html;

public static partial class HtmlComputedStyleEngine {
    /// <summary>
    /// Evaluates whether a CSS media query list applies to the requested OfficeIMO media context.
    /// </summary>
    public static bool IsApplicableMedia(string mediaText, HtmlCssMediaContext mediaContext) =>
        IsApplicableMedia(mediaText, MediaEnvironment.CreateDefault(mediaContext));

    internal static bool IsApplicableMedia(
        string mediaText,
        HtmlCssMediaContext mediaContext,
        HtmlRenderMediaFeatures mediaFeatures) {
        if (mediaFeatures == null) throw new ArgumentNullException(nameof(mediaFeatures));
        mediaFeatures.Validate();
        return IsApplicableMedia(mediaText, MediaEnvironment.CreateDefault(mediaContext, mediaFeatures));
    }

    /// <summary>
    /// Evaluates whether a CSS media query list applies to a media context and surface size.
    /// </summary>
    public static bool IsApplicableMedia(string mediaText, HtmlCssMediaContext mediaContext, double surfaceWidth, double surfaceHeight) {
        return IsApplicableMedia(mediaText, mediaContext, surfaceWidth, surfaceHeight, new HtmlRenderMediaFeatures());
    }

    /// <summary>
    /// Evaluates whether a CSS media query list applies to a media context, surface size,
    /// and deterministic static device environment.
    /// </summary>
    public static bool IsApplicableMedia(
        string mediaText,
        HtmlCssMediaContext mediaContext,
        double surfaceWidth,
        double surfaceHeight,
        HtmlRenderMediaFeatures mediaFeatures) {
        if (surfaceWidth <= 0D || double.IsNaN(surfaceWidth) || double.IsInfinity(surfaceWidth)) {
            throw new ArgumentOutOfRangeException(nameof(surfaceWidth));
        }
        if (surfaceHeight <= 0D || double.IsNaN(surfaceHeight) || double.IsInfinity(surfaceHeight)) {
            throw new ArgumentOutOfRangeException(nameof(surfaceHeight));
        }
        if (mediaFeatures == null) throw new ArgumentNullException(nameof(mediaFeatures));
        mediaFeatures.Validate();

        return IsApplicableMedia(mediaText, new MediaEnvironment(mediaContext, surfaceWidth, surfaceHeight, mediaFeatures));
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
        bool anyBranchApplies = false;
        foreach (string branch in SplitTopLevelMediaOrBranches(mediaQuery)) {
            if (branch.Length == 0) return false;
            if (AreMediaFeaturesApplicable(branch, environment)
                && (ContainsMediaType(branch, "all") || ContainsMediaType(branch, activeType) || !ContainsExplicitMediaType(branch))) {
                anyBranchApplies = true;
            }
        }

        return anyBranchApplies;
    }

    private static IEnumerable<string> SplitTopLevelMediaOrBranches(string mediaQuery) {
        int depth = 0;
        char quote = '\0';
        int start = 0;
        for (int index = 0; index < mediaQuery.Length; index++) {
            char current = mediaQuery[index];
            if (quote != '\0') {
                if (current == quote && !IsEscaped(mediaQuery, index)) quote = '\0';
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
            if (current == ')') {
                if (depth > 0) depth--;
                continue;
            }
            if (depth != 0 || !IsMediaLogicalOperatorAt(mediaQuery, index, "or")) continue;

            yield return mediaQuery.Substring(start, index - start).Trim();
            index++;
            start = index + 1;
        }

        yield return mediaQuery.Substring(start).Trim();
    }

    private static bool IsMediaLogicalOperatorAt(string mediaQuery, int index, string logicalOperator) {
        if (index < 0
            || index + logicalOperator.Length > mediaQuery.Length
            || !string.Equals(mediaQuery.Substring(index, logicalOperator.Length), logicalOperator, StringComparison.OrdinalIgnoreCase)
            || IsEscaped(mediaQuery, index)) {
            return false;
        }

        int before = index - 1;
        int after = index + logicalOperator.Length;
        return (before < 0 || !IsIdentifierCharacter(mediaQuery[before]))
            && (after >= mediaQuery.Length || !IsIdentifierCharacter(mediaQuery[after]));
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

        if (TryEvaluateIntegerFeature(feature, "color", environment.Features.ColorBitsPerComponent, out bool colorApplies)) {
            return colorApplies;
        }
        if (TryEvaluateIntegerFeature(feature, "monochrome", environment.Features.MonochromeBitsPerPixel, out bool monochromeApplies)) {
            return monochromeApplies;
        }

        if (TryEvaluateMediaLengthFeature(feature, environment, out bool lengthApplies)) return lengthApplies;

        if (feature.StartsWith("orientation", StringComparison.Ordinal)) {
            int colon = feature.IndexOf(':');
            if (colon < 0) return false;
            string value = feature.Substring(colon + 1).Trim();
            bool landscape = environment.Width >= environment.Height;
            return string.Equals(value, landscape ? "landscape" : "portrait", StringComparison.Ordinal);
        }

        if (TryEvaluateResolutionFeature(feature, environment.Features.ResolutionDpi, out bool resolutionApplies)) {
            return resolutionApplies;
        }
        if (feature == "prefers-color-scheme") {
            return true;
        }
        if (feature == "prefers-reduced-motion") {
            return environment.Features.ReducedMotion == HtmlReducedMotionPreference.Reduce;
        }
        if (TryReadMediaFeatureValue(feature, "prefers-color-scheme", out string colorScheme)) {
            return string.Equals(
                colorScheme,
                environment.Features.PreferredColorScheme == HtmlPreferredColorScheme.Dark ? "dark" : "light",
                StringComparison.Ordinal);
        }
        if (TryReadMediaFeatureValue(feature, "prefers-reduced-motion", out string motion)) {
            return string.Equals(
                motion,
                environment.Features.ReducedMotion == HtmlReducedMotionPreference.Reduce ? "reduce" : "no-preference",
                StringComparison.Ordinal);
        }
        if (feature == "pointer") return environment.Features.Pointer != HtmlPointerCapability.None;
        if (feature == "any-pointer") return environment.Features.AnyPointer != HtmlPointerCapability.None;
        if (feature == "hover") return environment.Features.Hover != HtmlHoverCapability.None;
        if (feature == "any-hover") return environment.Features.AnyHover != HtmlHoverCapability.None;
        if (TryReadMediaFeatureValue(feature, "pointer", out string pointer)) {
            return string.Equals(pointer, PointerValue(environment.Features.Pointer), StringComparison.Ordinal);
        }
        if (TryReadMediaFeatureValue(feature, "any-pointer", out string anyPointer)) {
            return string.Equals(anyPointer, PointerValue(environment.Features.AnyPointer), StringComparison.Ordinal);
        }
        if (TryReadMediaFeatureValue(feature, "hover", out string hover)) {
            return string.Equals(hover, HoverValue(environment.Features.Hover), StringComparison.Ordinal);
        }
        if (TryReadMediaFeatureValue(feature, "any-hover", out string anyHover)) {
            return string.Equals(anyHover, HoverValue(environment.Features.AnyHover), StringComparison.Ordinal);
        }
        if (TryReadMediaFeatureValue(feature, "scripting", out string scripting)) {
            return string.Equals(scripting, "none", StringComparison.Ordinal);
        }
        if (TryReadMediaFeatureValue(feature, "update", out string update)) {
            return string.Equals(update, "none", StringComparison.Ordinal);
        }
        if (TryReadMediaFeatureValue(feature, "overflow-block", out string overflowBlock)) {
            string expected = environment.Context == HtmlCssMediaContext.Print ? "paged" : "scroll";
            return string.Equals(overflowBlock, expected, StringComparison.Ordinal);
        }
        if (TryReadMediaFeatureValue(feature, "overflow-inline", out string overflowInline)) {
            return string.Equals(overflowInline, environment.Context == HtmlCssMediaContext.Print ? "none" : "scroll", StringComparison.Ordinal);
        }

        return false;
    }

    private static bool TryEvaluateIntegerFeature(string feature, string name, int actual, out bool applies) {
        applies = false;
        int colon = feature.IndexOf(':');
        string featureName = (colon < 0 ? feature : feature.Substring(0, colon)).Trim();
        bool recognized = featureName == name || featureName == "min-" + name || featureName == "max-" + name;
        if (!recognized) return false;
        if (colon < 0) {
            applies = actual > 0;
            return true;
        }
        if (!int.TryParse(feature.Substring(colon + 1).Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out int expected)
            || expected < 0) {
            return true;
        }
        applies = featureName.StartsWith("min-", StringComparison.Ordinal)
            ? actual >= expected
            : featureName.StartsWith("max-", StringComparison.Ordinal)
                ? actual <= expected
                : actual == expected;
        return true;
    }

    private static bool TryEvaluateResolutionFeature(string feature, double actualDpi, out bool applies) {
        applies = false;
        int colon = feature.IndexOf(':');
        string name = (colon < 0 ? feature : feature.Substring(0, colon)).Trim();
        if (name != "resolution" && name != "min-resolution" && name != "max-resolution") return false;
        if (colon < 0) {
            applies = actualDpi > 0D;
            return true;
        }
        string value = feature.Substring(colon + 1).Trim();
        int unitStart = value.Length;
        while (unitStart > 0 && char.IsLetter(value[unitStart - 1])) unitStart--;
        string number = value.Substring(0, unitStart).Trim();
        if (number.Length == 0
            || !double.TryParse(number, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed)
            || parsed < 0D
            || double.IsNaN(parsed)
            || double.IsInfinity(parsed)) {
            return true;
        }
        string unit = value.Substring(unitStart).Trim();
        double expectedDpi;
        switch (unit) {
            case "dpi": expectedDpi = parsed; break;
            case "dpcm": expectedDpi = parsed * 2.54D; break;
            case "dppx":
            case "x": expectedDpi = parsed * HtmlRenderOptions.CssPixelsPerInch; break;
            default: return true;
        }
        applies = name.StartsWith("min-", StringComparison.Ordinal)
            ? actualDpi >= expectedDpi
            : name.StartsWith("max-", StringComparison.Ordinal)
                ? actualDpi <= expectedDpi
                : Math.Abs(actualDpi - expectedDpi) <= 0.000001D;
        return true;
    }

    private static bool TryReadMediaFeatureValue(string feature, string name, out string value) {
        value = string.Empty;
        int colon = feature.IndexOf(':');
        if (colon < 0 || !string.Equals(feature.Substring(0, colon).Trim(), name, StringComparison.Ordinal)) return false;
        value = feature.Substring(colon + 1).Trim();
        return value.Length > 0;
    }

    private static string PointerValue(HtmlPointerCapability capability) {
        if (capability == HtmlPointerCapability.Coarse) return "coarse";
        if (capability == HtmlPointerCapability.Fine) return "fine";
        return "none";
    }

    private static string HoverValue(HtmlHoverCapability capability) =>
        capability == HtmlHoverCapability.Hover ? "hover" : "none";

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
        if (cursor < value.Length && (value[cursor] == 'e' || value[cursor] == 'E')) {
            cursor++;
            if (cursor < value.Length && (value[cursor] == '+' || value[cursor] == '-')) cursor++;
            int exponentStart = cursor;
            while (cursor < value.Length && char.IsDigit(value[cursor])) cursor++;
            if (cursor == exponentStart) return false;
        }

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
        internal MediaEnvironment(
            HtmlCssMediaContext context,
            double width,
            double height,
            HtmlRenderMediaFeatures? features = null) {
            Context = context;
            Width = width;
            Height = height;
            Features = features ?? new HtmlRenderMediaFeatures();
        }

        internal HtmlCssMediaContext Context { get; }
        internal double Width { get; }
        internal double Height { get; }
        internal HtmlRenderMediaFeatures Features { get; }

        internal static MediaEnvironment CreateDefault(
            HtmlCssMediaContext context,
            HtmlRenderMediaFeatures? features = null) {
            if (context == HtmlCssMediaContext.Print) {
                return new MediaEnvironment(
                    context,
                    OfficePageSizes.A4.WidthInches * HtmlRenderOptions.CssPixelsPerInch,
                    OfficePageSizes.A4.HeightInches * HtmlRenderOptions.CssPixelsPerInch,
                    features);
            }

            return new MediaEnvironment(context, 816D, 1056D, features);
        }
    }
}
