using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderStyleResolver {
    private double ResolveComputedLineHeight(HtmlComputedStyle computed, double fontSize, HtmlRenderBoxStyle? parent,
        out double alignmentLineHeight, out bool usesNaturalLineHeight) {
        string value = computed.GetValue("line-height").Trim();
        usesNaturalLineHeight = UsesNaturalLineHeight(value);
        // Numbers and normal remain font-relative; explicit lengths and percentages
        // inherit their computed value, including across more than one font change.
        if (computed.IsInheritedValue("line-height") && parent != null
            && value.Length > 0 && !value.Equals("normal", StringComparison.OrdinalIgnoreCase)
            && !double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out _)) {
            alignmentLineHeight = parent.InlineAlignmentLineHeight
                ?? HtmlRenderCssValues.ResolveInlineAlignmentLineHeight(parent.LineHeight, parent.Font.Size);
            return parent.LineHeight;
        }
        double lineHeight = ResolveLineHeight(value, fontSize, out double? multiplier);
        alignmentLineHeight = HtmlRenderCssValues.ResolveInlineAlignmentLineHeight(lineHeight, fontSize, multiplier);
        return lineHeight;
    }

    private bool UsesNaturalLineHeight(string value) => !_options.HasExplicitDefaultLineHeight
        && (value.Length == 0 || value.Equals("normal", StringComparison.OrdinalIgnoreCase));

    private double ResolveLineHeight(string value, double fontSize, out double? multiplier) {
        multiplier = _options.DefaultLineHeight;
        if (value.Length == 0 || value.Equals("normal", StringComparison.OrdinalIgnoreCase)) return fontSize * multiplier.Value;
        if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed) && parsed >= 0D) {
            multiplier = parsed;
            return fontSize * parsed;
        }
        if (TryResolveLength(value, fontSize, fontSize, _rootFontSize, out double lineHeight, out bool calculated)
            && (lineHeight >= 0D || calculated)) {
            multiplier = null;
            return Math.Max(0D, lineHeight);
        }
        return fontSize * multiplier.Value;
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
