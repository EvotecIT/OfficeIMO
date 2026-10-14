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
        "baseline" => inherited,
        _ => 0
    };

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
