using OfficeIMO.Html.Css;

namespace OfficeIMO.Html;

/// <summary>Shared Browser size/provenance policy; numeric evaluation remains with the owned CSS values.</summary>
internal static class HtmlFontSizePolicy {
    internal static HtmlFontSizeContext Resolve(
        HtmlComputedStyle style,
        HtmlFontSizeContext inherited,
        string cssFamily,
        HtmlRenderingFontDefaults defaults,
        HtmlFontSizeContext htmlDefault,
        HtmlCssLengthResolutionContext lengthContext) {
        bool monospace = IsSoleGenericMonospace(cssFamily);
        HtmlFontSizeContext resolved;
        if (style.IsOriginRevertedValue("font-size")) {
            resolved = htmlDefault;
        } else if (style.IsResetValue("font-size")) {
            resolved = new HtmlFontSizeContext(defaults.Medium(monospace), false, "medium", inherited.SoleGenericMonospace);
        } else if (style.IsImplicitlyInheritedValue("font-size")) {
            resolved = htmlDefault;
        } else if (style.IsInheritedValue("font-size")) {
            resolved = inherited;
        } else {
            string value = style.GetValue("font-size").Trim().ToLowerInvariant();
            if (defaults.TryResolveKeyword(value, inherited.Pixels, monospace, out double keyword)) {
                bool relativeKeyword = value is "smaller" or "larger";
                resolved = new HtmlFontSizeContext(keyword, relativeKeyword && inherited.IsAbsolute,
                    relativeKeyword ? string.Empty : value, inherited.SoleGenericMonospace);
            } else if (HtmlRenderCssValues.TryLength(value, lengthContext, out double pixels, out _, out HtmlCssMathExpression? expression)
                && !double.IsNaN(pixels) && !double.IsInfinity(pixels) && pixels >= 0D) {
                resolved = new HtmlFontSizeContext(pixels, IsAbsoluteLength(expression!, inherited.IsAbsolute),
                    string.Empty, inherited.SoleGenericMonospace);
            } else {
                resolved = htmlDefault;
            }
        }
        return ApplyGenericTransition(resolved, monospace, defaults);
    }

    /// <summary>Identifies CSS generic syntax before a physical face or quoted family is resolved.</summary>
    internal static bool IsSoleGenericMonospace(string family) {
        bool found = false;
        foreach (HtmlCssToken token in HtmlCssTokenizer.Enumerate(family)) {
            if (token.Kind is HtmlCssTokenKind.Whitespace or HtmlCssTokenKind.Comment or HtmlCssTokenKind.EndOfFile) continue;
            if (found || token.Kind != HtmlCssTokenKind.Identifier
                || !string.Equals(token.Value, "monospace", StringComparison.OrdinalIgnoreCase)) return false;
            found = true;
        }
        return found;
    }

    /// <summary>Preserves the existing HTML tag scales while clearing an inherited keyword when a relative UA rule applies.</summary>
    internal static HtmlFontSizeContext ResolveHtmlDefault(string tag, HtmlFontSizeContext inherited, bool browser = false) => tag switch {
        "h1" => inherited.WithRelativeSize(inherited.Pixels * 2D),
        "h2" => inherited.WithRelativeSize(inherited.Pixels * 1.5D),
        "h3" => inherited.WithRelativeSize(inherited.Pixels * 1.17D),
        "h4" => inherited.WithRelativeSize(inherited.Pixels),
        "h5" => inherited.WithRelativeSize(inherited.Pixels * 0.83D),
        "h6" => inherited.WithRelativeSize(inherited.Pixels * 0.67D),
        "small" => inherited.WithRelativeSize(browser ? inherited.Pixels / 1.2D : inherited.Pixels * 0.83D),
        "sup" or "sub" => inherited.WithRelativeSize(inherited.Pixels / 1.2D),
        "big" => inherited.WithRelativeSize(inherited.Pixels * (browser ? 1.2D : 1.17D)),
        _ => inherited
    };

    private static bool IsAbsoluteLength(HtmlCssMathExpression expression, bool inheritedAbsolute) {
        // The typed root distinguishes calculated values and literal font-relative
        // units. No declaration substrings or second expression traversal are needed.
        if (inheritedAbsolute || expression.IsCalculated) return true;
        if (expression.Type == HtmlCssNumericType.Percentage) return false;
        return expression.Unit is not (HtmlCssLengthUnit.Em or HtmlCssLengthUnit.Ex or HtmlCssLengthUnit.Ch);
    }

    private static HtmlFontSizeContext ApplyGenericTransition(
        HtmlFontSizeContext value,
        bool monospace,
        HtmlRenderingFontDefaults defaults) {
        double pixels = value.Pixels;
        if (!value.IsAbsolute && monospace != value.SoleGenericMonospace) {
            if (value.AbsoluteKeyword.Length > 0
                && defaults.TryResolveKeyword(value.AbsoluteKeyword, pixels, monospace, out double keyword)) {
                pixels = keyword;
            } else {
                double scale = defaults.FixedMedium / defaults.StandardMedium;
                pixels = monospace ? pixels * scale : pixels / scale;
            }
        }
        return new HtmlFontSizeContext(pixels, value.IsAbsolute, value.AbsoluteKeyword, monospace);
    }
}
