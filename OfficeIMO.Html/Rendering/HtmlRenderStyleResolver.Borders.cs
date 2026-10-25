using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderStyleResolver {
    private void ApplyBorderAndOutlinePaint(HtmlComputedStyle computed, double reference, double fontSize, HtmlRenderBoxStyle style) {
        ResolvedBoxStroke stroke = ResolveBoxStroke(computed, reference, fontSize, style.Color);
        style.BorderDeclared = stroke.BorderDeclared;
        if (stroke.BorderSupported) {
            style.Borders = stroke.Borders;
        } else {
            style.Borders = HtmlRenderBorderEdges.Uniform(0D, "none", style.Color);
            style.UnsupportedBorderPaint = stroke.BorderDetail;
        }

        style.BorderRadius = NormalizeCssValue(computed.GetValue("border-radius"), "0");
        style.BorderTopLeftRadius = computed.GetValue("border-top-left-radius").Trim();
        style.BorderTopRightRadius = computed.GetValue("border-top-right-radius").Trim();
        style.BorderBottomRightRadius = computed.GetValue("border-bottom-right-radius").Trim();
        style.BorderBottomLeftRadius = computed.GetValue("border-bottom-left-radius").Trim();

        if (stroke.OutlineSupported) {
            style.OutlineWidth = stroke.OutlineWidth;
            style.OutlineStyle = stroke.OutlineStyle;
            style.OutlineColor = stroke.OutlineColor;
            style.OutlineColorInvert = stroke.OutlineColorInvert;
            style.OutlineOffset = stroke.OutlineOffset;
        } else {
            style.OutlineWidth = 0D;
            style.OutlineStyle = "none";
            style.UnsupportedOutlinePaint = stroke.OutlineDetail;
        }
    }
}
