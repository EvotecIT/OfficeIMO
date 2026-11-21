namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderStyleResolver {
    private static string ResolveTextVerticalAlign(string tag, HtmlComputedStyle computed) {
        // HTML's script alignment is a user-agent default. An authored baseline or
        // initial/unset reset overrides it independently of the element's font size.
        if (!computed.IsOriginRevertedValue("vertical-align")) {
            if (computed.IsResetValue("vertical-align")) return "baseline";
            string value = computed.GetValue("vertical-align").Trim().ToLowerInvariant();
            if (value.Length > 0) return value;
        }
        return tag == "sup" ? "super" : tag == "sub" ? "sub" : "baseline";
    }

    private static int ResolveTextBaselineLevel(string verticalAlign, int inherited) => verticalAlign switch {
        "super" => inherited + 1,
        "sub" => inherited - 1,
        "baseline" or "middle" or "text-top" or "text-bottom" or "top" or "bottom" => inherited,
        _ => 0
    };

    // CSS 2.1 10.8.1 for an inline text box: its box is the line-height box, with the
    // font's ascent at 0.8em and descent at 0.2em. The parent's x-height is 0.5em.
    private static double ResolveVerticalAlignKeywordShift(string verticalAlign, double parentFontSize, double fontSize, double lineHeight) {
        double ascent = lineHeight / 2D + fontSize * 0.3D;
        double descent = lineHeight - ascent;
        return verticalAlign switch {
            "middle" => ascent - lineHeight / 2D - parentFontSize * 0.25D,
            "text-top" => ascent - parentFontSize * 0.8D,
            "text-bottom" => parentFontSize * 0.2D - descent,
            _ => 0D
        };
    }

    private double ResolveTextBaselineOffset(
        string verticalAlign,
        double inheritedOffset,
        double parentFontSize,
        double fontSize,
        double lineHeight) {
        // Font size is already resolved, including the sup/sub default. CSS
        // vertical-align changes displacement, never the measurement or paint font.
        if (verticalAlign == "super") return inheritedOffset - parentFontSize * 0.30D;
        if (verticalAlign == "sub") return inheritedOffset + parentFontSize * 0.15D;
        if (verticalAlign == "baseline") return inheritedOffset;
        return TryResolveLength(verticalAlign, lineHeight, fontSize, _rootFontSize, out double shift)
            ? inheritedOffset - shift
            : inheritedOffset;
    }
}
