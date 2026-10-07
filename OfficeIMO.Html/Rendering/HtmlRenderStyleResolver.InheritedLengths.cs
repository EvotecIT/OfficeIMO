using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderStyleResolver {
    private double ResolveInheritedLineHeight(HtmlComputedStyle computed, double fontSize, HtmlRenderBoxStyle? parent) {
        string value = computed.GetValue("line-height").Trim();
        // Numbers and normal remain font-relative; explicit lengths and percentages
        // inherit their computed value, including across more than one font change.
        if (computed.IsInheritedValue("line-height") && parent != null
            && value.Length > 0 && !value.Equals("normal", StringComparison.OrdinalIgnoreCase)
            && !double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out _)) return parent.LineHeight;
        return ResolveLineHeight(value, fontSize);
    }

    private static void ApplyInheritedTextShadows(HtmlComputedStyle computed, HtmlRenderBoxStyle style, HtmlRenderBoxStyle? parent) {
        if (parent == null || !computed.IsInheritedValue("text-shadow")) return;
        style.UnsupportedTextShadow = parent.UnsupportedTextShadow;
        style.TextShadowLayerCount = parent.TextShadowLayerCount;
        style.TextShadows = parent.TextShadows.Select(shadow => shadow.UsesCurrentColor
            ? new HtmlCssTextShadow(OfficeColor.FromRgb(style.Color.R, style.Color.G, style.Color.B), style.Color.A / 255D,
                shadow.OffsetX, shadow.OffsetY, shadow.BlurRadius, usesCurrentColor: true)
            : shadow).ToArray();
    }
}
