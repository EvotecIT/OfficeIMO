using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private void ResolveInlineTextVerticalPlacement(
        InlineSegment segment,
        bool hasReplacedImage,
        double lineY,
        double lineHeight,
        double baseline,
        out double textY,
        out double paintHeight,
        out double paintTopOverflow) {
        HtmlRenderBoxStyle style = segment.Run.Style;
        double fontSize = style.Font.Size;
        textY = hasReplacedImage
            ? lineY + Math.Max(0D, baseline - ResolveTextAscent(style))
            : lineY + Math.Min(0D, (lineHeight - fontSize) / 2D);
        paintHeight = Math.Max(lineHeight, fontSize);
        paintTopOverflow = 0D;
        if (hasReplacedImage || lineHeight >= fontSize) return;

        HtmlTextFaceMetrics? face = ResolveTextFaceMetrics(segment.Text, style);
        if (!face.HasValue || face.Value.Height <= 0D
            || double.IsNaN(face.Value.Height) || double.IsInfinity(face.Value.Height)
            || double.IsNaN(face.Value.BaselineOffset) || double.IsInfinity(face.Value.BaselineOffset)
            || face.Value.BaselineOffset < 0D || face.Value.BaselineOffset > face.Value.Height) return;

        // CSS centers the face's ascent/descent box in the used line height. Positioned Drawing
        // text anchors its baseline one em below Y, so translate that browser baseline back to Y.
        textY = lineY + (lineHeight - face.Value.Height) / 2D
            + face.Value.BaselineOffset - fontSize;
        paintHeight = Math.Max(lineHeight, face.Value.Height);
        paintTopOverflow = Math.Max(0D, face.Value.BaselineOffset - fontSize);
    }

    private HtmlTextFaceMetrics? ResolveTextFaceMetrics(string text, HtmlRenderBoxStyle style) {
        if (_fonts.TryResolveFaceForText(text, style.Font.FamilyName, style.FontDescriptor, out OfficeFontFace? face)
            && face?.Program is IOfficeFontBaselineMetrics baseline) {
            return new HtmlTextFaceMetrics(face.Program.LineHeight(style.Font.Size),
                baseline.BaselineOffset(style.Font.Size));
        }
        return _options.FallbackTextFaceMetrics?.Invoke(text, style.Font, style.FontDescriptor);
    }
}
