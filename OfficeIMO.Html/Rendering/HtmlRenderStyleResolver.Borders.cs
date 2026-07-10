using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderStyleResolver {
    private void ApplyBorderAndOutlinePaint(HtmlComputedStyle computed, double reference, double fontSize, HtmlRenderBoxStyle style) {
        style.BorderDeclared = HtmlCssBoxStrokeParser.HasBorderDeclaration(computed);
        if (HtmlCssBoxStrokeParser.TryParseBorder(
                computed,
                reference,
                fontSize,
                _options.DefaultFontSize,
                style.Color,
                out double borderWidth,
                out string borderStyle,
                out OfficeColor borderColor,
                out string borderDetail)) {
            style.BorderWidth = borderWidth;
            style.BorderStyle = borderStyle;
            style.BorderColor = borderColor;
        } else {
            style.BorderWidth = 0D;
            style.BorderStyle = "none";
            style.UnsupportedBorderPaint = borderDetail;
        }

        style.BorderRadius = NormalizeCssValue(computed.GetValue("border-radius"), "0");
        style.BorderTopLeftRadius = computed.GetValue("border-top-left-radius").Trim();
        style.BorderTopRightRadius = computed.GetValue("border-top-right-radius").Trim();
        style.BorderBottomRightRadius = computed.GetValue("border-bottom-right-radius").Trim();
        style.BorderBottomLeftRadius = computed.GetValue("border-bottom-left-radius").Trim();

        if (HtmlCssBoxStrokeParser.TryParseOutline(
                computed,
                reference,
                fontSize,
                _options.DefaultFontSize,
                style.Color,
                out double outlineWidth,
                out string outlineStyle,
                out OfficeColor outlineColor,
                out double outlineOffset,
                out string outlineDetail)) {
            style.OutlineWidth = outlineWidth;
            style.OutlineStyle = outlineStyle;
            style.OutlineColor = outlineColor;
            style.OutlineOffset = outlineOffset;
        } else {
            style.OutlineWidth = 0D;
            style.OutlineStyle = "none";
            style.UnsupportedOutlinePaint = outlineDetail;
        }
    }
}
