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

    /// <summary>Resolves generated text through the same CSS leading owner as ordinary inline text.</summary>
    internal double ResolveGeneratedLineHeight(string value, OfficeFontInfo font, double viewportWidth, double viewportHeight,
        out double alignmentLineHeight, out bool usesNaturalLineHeight) {
        value = value.Trim();
        usesNaturalLineHeight = UsesNaturalLineHeight(value);
        double previousWidth = _activeContainerWidth, previousHeight = _activeContainerHeight;
        double previousViewportWidth = _viewportWidth, previousViewportHeight = _viewportHeight;
        double previousAdvance = _activeCharacterAdvance;
        bool previousUpright = _activeUprightVerticalText;
        try {
            // Page-margin text has no ancestor query container or body writing mode.
            // Its viewport fallback uses the print paper and its own selected face.
            _viewportWidth = _activeContainerWidth = viewportWidth;
            _viewportHeight = _activeContainerHeight = viewportHeight;
            _activeUprightVerticalText = false;
            _activeCharacterAdvance = ContainsCharacterUnit(value)
                ? _measureCharacterAdvance(new HtmlRenderBoxStyle {
                    Font = font, FontDescriptor = OfficeFontFaceDescriptor.FromStyle(font.Style)
                }) : double.NaN;
            double lineHeight = ResolveLineHeight(value, font.Size, out double? multiplier);
            alignmentLineHeight = HtmlRenderCssValues.ResolveInlineAlignmentLineHeight(lineHeight, font.Size, multiplier);
            return lineHeight;
        } finally {
            _viewportWidth = previousViewportWidth;
            _viewportHeight = previousViewportHeight;
            _activeContainerWidth = previousWidth;
            _activeContainerHeight = previousHeight;
            _activeCharacterAdvance = previousAdvance;
            _activeUprightVerticalText = previousUpright;
        }
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
        if (TryResolveLength(value, fontSize, fontSize, _rootFontSize, out double lineHeight) && lineHeight >= 0D) {
            multiplier = null;
            return lineHeight;
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
