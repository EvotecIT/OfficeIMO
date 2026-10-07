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
        // This frame starts one source em above the baseline, rather than at the
        // face ascent. Its lower edge must include the selected face's descent.
        paintHeight = Math.Max(lineHeight, fontSize + face.Value.Height - face.Value.BaselineOffset);
        paintTopOverflow = Math.Max(0D, face.Value.BaselineOffset - fontSize);
    }

    private double ResolveInlineSharedBaseline(InlineLine line, HtmlRenderBoxStyle paragraphStyle, ref double lineHeight) {
        bool hasImage = line.HasReplacedImage;
        string strutText = line.Segments.FirstOrDefault(segment => segment.Run.AtomicBlock == null
            && segment.Text.Length > 0)?.Text ?? string.Empty;
        double baseline = ResolveInlineTextBaseline(strutText, paragraphStyle, hasImage);
        double descent = Math.Max(0D, paragraphStyle.LineHeight - baseline);
        foreach (InlineSegment segment in line.Segments) {
            HtmlRenderFlowBlock? atomic = segment.Run.AtomicBlock;
            if (atomic != null && !hasImage) continue;
            double sourceBaseline = atomic != null
                ? Math.Min(atomic.Height, Math.Max(0D, segment.Run.AtomicBaseline ?? atomic.Height))
                : ResolveInlineTextBaseline(segment.Text, segment.Run.Style, hasImage);
            baseline = Math.Max(baseline, sourceBaseline);
            descent = Math.Max(descent, (atomic?.Height ?? segment.Run.Style.LineHeight) - sourceBaseline);
        }
        // Replaced images occupy flow space through their baseline. When a
        // selected short-line face raises that baseline, advance the line by the
        // same ascent/descent so a following block or legal cut cannot cross it.
        if (hasImage) lineHeight = Math.Max(lineHeight, baseline + descent);
        return baseline;
    }

    private double ResolveInlineTextBaseline(string text, HtmlRenderBoxStyle style, bool hasImage) {
        HtmlTextFaceMetrics? face = ResolveInlineShortLineFace(text, style);
        // Alignment and ink overflow are separate. The paragraph strut and each
        // text run use the same selected face, including beside an actual image.
        return face.HasValue
            ? (style.LineHeight - face.Value.Height) / 2D + face.Value.BaselineOffset
            : hasImage ? ResolveTextAscent(style) : style.Font.Size;
    }

    private HtmlTextFaceMetrics? ResolveInlineShortLineFace(InlineSegment segment) =>
        ResolveInlineShortLineFace(segment.Text, segment.Run.Style);

    private HtmlTextFaceMetrics? ResolveInlineShortLineFace(string text, HtmlRenderBoxStyle style) {
        if (style.LineHeight >= style.Font.Size) return null;
        HtmlTextFaceMetrics? face = ResolveTextFaceMetrics(text, style);
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
