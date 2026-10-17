using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private void ResolveInlineTextVerticalPlacement(
        InlineSegment segment,
        bool alignBaseline,
        double lineY,
        double lineHeight,
        double baseline,
        bool useOrdinaryFaceMetrics,
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
        HtmlTextFaceMetrics? face = useOrdinaryFaceMetrics && style.Font.Size > 0D
            ? ResolveTextFaceMetrics(segment.Text, style)
            : ResolveInlineShortLineFace(segment);
        if (!IsUsableTextFaceMetrics(face)) return;
        if (!face.HasValue) return;

        // CSS centers the face's ascent/descent box in the used line height. Positioned Drawing
        // text anchors its baseline one em below Y, so translate that browser baseline back to Y.
        if (!alignBaseline) {
            textY = useOrdinaryFaceMetrics && TryResolveNormalInlineTextMetrics(style, face, out InlineTextLineMetrics metrics)
                ? lineY + metrics.Baseline - fontSize
                : lineY + (lineHeight - face.Value.Height) / 2D
                    + face.Value.BaselineOffset - fontSize;
        }
        // This frame starts one source em above the baseline, rather than at the
        // face ascent. Its lower edge must include the selected face's descent.
        paintHeight = Math.Max(lineHeight, fontSize + face.Value.Height - face.Value.BaselineOffset);
        paintTopOverflow = Math.Max(0D, face.Value.BaselineOffset - fontSize);
    }

    private double ResolveInlineSharedBaseline(InlineLine line, HtmlRenderBoxStyle paragraphStyle, ref double lineHeight) {
        bool hasImage = line.HasReplacedImage;
        bool useNormalImageMetrics = UsesNormalImageLineMetrics(line, paragraphStyle);
        string strutText = line.Segments.FirstOrDefault(segment => segment.Run.AtomicBlock == null
            && segment.Text.Length > 0)?.Text ?? string.Empty;
        InlineTextLineMetrics strut = ResolveInlineSharedTextMetrics(strutText, paragraphStyle, hasImage, useNormalImageMetrics);
        double baseline = strut.Baseline;
        double descent = Math.Max(0D, strut.Height - baseline);
        foreach (InlineSegment segment in line.Segments) {
            if (!InlineLine.ContributesFlowContent(segment.Run)) continue;
            HtmlRenderFlowBlock? atomic = segment.Run.AtomicBlock;
            if (atomic != null && !hasImage) continue;
            InlineTextLineMetrics source = atomic != null
                ? new InlineTextLineMetrics(atomic.Height,
                    Math.Min(atomic.Height, Math.Max(0D, segment.Run.AtomicBaseline ?? atomic.Height)))
                : ResolveInlineSharedTextMetrics(segment.Text, segment.Run.Style, hasImage, useNormalImageMetrics);
            double sourceBaseline = source.Baseline;
            baseline = Math.Max(baseline, sourceBaseline);
            descent = Math.Max(descent, source.Height - sourceBaseline);
        }
        // Replace the preliminary em-based estimate with the same used metrics
        // that align the image. Keeping its old descent as a minimum would leave
        // text-containing image lines taller than their image-only equivalents.
        if (hasImage) lineHeight = useNormalImageMetrics
            ? Math.Max(0.01D, baseline + descent)
            : Math.Max(lineHeight, baseline + descent);
        return baseline;
    }

    private double ResolveInlineTextBaseline(string text, HtmlRenderBoxStyle style, bool hasImage) {
        if (!hasImage && style.Font.Size > 0D
            && TryResolveNormalInlineTextMetrics(style, ResolveTextFaceMetrics(text, style), out InlineTextLineMetrics normal))
            return normal.Baseline;
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
        return IsUsableTextFaceMetrics(face) ? face : null;
    }

    private static bool IsUsableTextFaceMetrics(HtmlTextFaceMetrics? face) =>
        face.HasValue && face.Value.Height > 0D && !double.IsNaN(face.Value.Height) && !double.IsInfinity(face.Value.Height)
        && !double.IsNaN(face.Value.BaselineOffset) && !double.IsInfinity(face.Value.BaselineOffset)
        && face.Value.BaselineOffset >= 0D && face.Value.BaselineOffset <= face.Value.Height;

    /// <summary>Resolves horizontal CSS leading from used font boxes while preserving the program's raw metrics and glyph ink.</summary>
    private static bool TryResolveNormalInlineTextMetrics(HtmlRenderBoxStyle style,
        HtmlTextFaceMetrics? face, out InlineTextLineMetrics metrics, bool allowShortLine = false) {
        metrics = default;
        if (style.Font.Size <= 0D || !allowShortLine && style.LineHeight < style.Font.Size
            || !TryResolveUsedInlineFontBox(face, out InlineTextLineMetrics fontBox)) return false;
        double ascent = fontBox.Baseline;
        double descent = fontBox.Height - ascent;
        double naturalAdvance = face!.Value.NaturalAdvance ?? face.Value.Height;
        if (double.IsNaN(naturalAdvance) || double.IsInfinity(naturalAdvance)) naturalAdvance = face.Value.Height;
        double usedHeight = style.UsesNaturalLineHeight
            ? fontBox.Height + System.Math.Floor(System.Math.Max(0D,
                naturalAdvance - face.Value.Height) + 0.5D)
            : style.InlineAlignmentLineHeight
                ?? HtmlRenderCssValues.ResolveInlineAlignmentLineHeight(style.LineHeight, style.Font.Size);
        double leadingUnits = (usedHeight - ascent - descent) * 64D;
        // Divide signed fixed units before flooring the ascent share to a whole
        // pixel; the remaining leading belongs below the alphabetic baseline.
        double ascentLeading = Math.Floor(Math.Truncate(leadingUnits / 2D) / 64D);
        metrics = new InlineTextLineMetrics(usedHeight, ascent + ascentLeading);
        return true;
    }

    private HtmlTextFaceMetrics? ResolveTextFaceMetrics(string text, HtmlRenderBoxStyle style) {
        IOfficeFontProgram? program = ResolveTextFaceProgram(text, style);
        if (program is IOfficeFontBaselineMetrics baseline) {
            return new HtmlTextFaceMetrics(program.LineHeight(style.Font.Size),
                baseline.BaselineOffset(style.Font.Size), program.LineSpacingRatio * style.Font.Size);
        }
        return _options.FallbackTextFaceMetrics?.Invoke(text, style.Font, style.FontDescriptor);
    }

    private IOfficeFontProgram? ResolveTextFaceProgram(string text, HtmlRenderBoxStyle style) =>
        text.Length == 0
            ? _fonts.ResolveForText(text, style.Font.FamilyName, style.FontDescriptor, style.Font.Size, out _)
            : _fonts.TryResolveFaceForText(text, style.Font.FamilyName, style.FontDescriptor, style.Font.Size, out OfficeFontFace? face)
                ? face?.Program : null;

    private void ResolveInlineAnchorTextVerticalBounds(
        InlineSegment segment,
        double lineY,
        double lineHeight,
        double textY,
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
        if (TryResolveNormalInlineTextMetrics(segment.Run.Style, face, out _)) {
            anchorY = textY + segment.Run.Style.Font.Size - Math.Floor(face.Value.BaselineOffset + 0.5D);
        }
    }

    private readonly struct InlineTextLineMetrics {
        internal InlineTextLineMetrics(double height, double baseline) {
            Height = height;
            Baseline = baseline;
        }

        internal double Height { get; }
        internal double Baseline { get; }
    }
}
