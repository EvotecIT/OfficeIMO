using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private void ResolveInlineTextVerticalPlacement(
        InlineSegment segment,
        bool alignBaseline,
        double lineY,
        double lineHeight,
        double baseline,
        out double textY,
        out double paintHeight,
        out double paintTopOverflow) {
        HtmlRenderBoxStyle style = segment.Run.Style;
        double fontSize = style.Font.Size;
        textY = alignBaseline
            // Positioned text paints its alphabetic baseline one source em below Y.
            // The line ascent determines the shared baseline, not the paint-frame origin.
            ? lineY + baseline - fontSize
            : lineY + Math.Min(0D, (lineHeight - fontSize) / 2D);
        paintHeight = Math.Max(lineHeight, fontSize);
        paintTopOverflow = 0D;
        HtmlTextFaceMetrics? face = ResolveInlineShortLineFace(segment);
        if (!face.HasValue) return;

        // CSS centers the face's ascent/descent box in the used line height. Positioned Drawing
        // text anchors its baseline one em below Y, so translate that browser baseline back to Y.
        if (!alignBaseline) {
            textY = lineY + (lineHeight - face.Value.Height) / 2D
                + face.Value.BaselineOffset - fontSize;
        }
        paintHeight = Math.Max(lineHeight, face.Value.Height);
        paintTopOverflow = Math.Max(0D, face.Value.BaselineOffset - fontSize);
    }

    private double ResolveInlineMixedTextBaseline(InlineLine line, HtmlRenderBoxStyle paragraphStyle) {
        double baseline = paragraphStyle.Font.Size;
        foreach (InlineSegment segment in line.Segments) {
            if (segment.Run.AtomicBlock != null) continue;
            HtmlRenderBoxStyle style = segment.Run.Style;
            HtmlTextFaceMetrics? face = ResolveInlineShortLineFace(segment);
            // Alignment and face overflow are separate. A short authored line
            // supplies the same centered face baseline whether text stands alone
            // or shares the line with ordinary text or a first-letter fragment.
            double sourceBaseline = face.HasValue
                ? (style.LineHeight - face.Value.Height) / 2D + face.Value.BaselineOffset
                : style.Font.Size;
            baseline = Math.Max(baseline, sourceBaseline);
        }
        return baseline;
    }

    private HtmlTextFaceMetrics? ResolveInlineShortLineFace(InlineSegment segment) {
        HtmlRenderBoxStyle style = segment.Run.Style;
        if (style.LineHeight >= style.Font.Size) return null;
        HtmlTextFaceMetrics? face = ResolveTextFaceMetrics(segment.Text, style);
        if (!face.HasValue || face.Value.Height <= 0D
            || double.IsNaN(face.Value.Height) || double.IsInfinity(face.Value.Height)
            || double.IsNaN(face.Value.BaselineOffset) || double.IsInfinity(face.Value.BaselineOffset)
            || face.Value.BaselineOffset < 0D || face.Value.BaselineOffset > face.Value.Height) return null;
        return face;
    }

    private HtmlTextFaceMetrics? ResolveTextFaceMetrics(string text, HtmlRenderBoxStyle style) {
        if (_fonts.TryResolveFaceForText(text, style.Font.FamilyName, style.FontDescriptor, style.Font.Size, out OfficeFontFace? face)
            && face?.Program is IOfficeFontBaselineMetrics baseline) {
            return new HtmlTextFaceMetrics(face.Program.LineHeight(style.Font.Size),
                baseline.BaselineOffset(style.Font.Size));
        }
        return _options.FallbackTextFaceMetrics?.Invoke(text, style.Font, style.FontDescriptor);
    }

    private void ResolveInlineAnchorTextVerticalBounds(
        InlineSegment segment,
        double lineY,
        double lineHeight,
        out double anchorY,
        out double anchorHeight) {
        anchorY = lineY;
        anchorHeight = lineHeight;
        HtmlTextFaceMetrics? face = ResolveTextFaceMetrics(segment.Text, segment.Run.Style);
        if (!face.HasValue || face.Value.Height <= 0D || double.IsNaN(face.Value.Height)
            || double.IsInfinity(face.Value.Height)) return;

        // An inline anchor's CSS box follows its font's ascent/descent box, not
        // the full line box. Extra leading belongs to the line around it.
        anchorHeight = face.Value.Height;
    }
}
