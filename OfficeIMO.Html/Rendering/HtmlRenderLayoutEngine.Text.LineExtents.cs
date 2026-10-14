namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    /// <summary>
    /// Combines the ascent and descent of text-only horizontal lines before paint
    /// or float placement. Font size alone is not a face's ascent, and inherited
    /// leading belongs equally above and below each run's font box.
    /// </summary>
    private bool TryResolveTextOnlyLineMetrics(InlineLine line, HtmlRenderBoxStyle paragraphStyle,
        ref double lineHeight, out double baseline) {
        baseline = 0D;
        if (!line.HasFlowContent || IsVerticalWritingMode(paragraphStyle.WritingMode)
            || line.Segments.Any(segment => segment.Run.AtomicBlock != null)
            || !line.HasMixedTextSizes(paragraphStyle)
                && !line.Segments.Any(segment => Math.Abs(segment.Run.Style.BaselineOffset
                    - paragraphStyle.BaselineOffset) > 0.000001D)) return false;

        string strutText = line.Segments.FirstOrDefault(segment => segment.Text.Length > 0)?.Text ?? string.Empty;
        double strutAscent = ResolveTextLineAscent(strutText, paragraphStyle);
        double ascent = strutAscent;
        double descent = paragraphStyle.LineHeight - strutAscent;
        foreach (InlineSegment segment in line.Segments) {
            HtmlInlineRun run = segment.Run;
            if (run.IsFlowMarker || run.RunningStringElement != null || run.RunningElementAssignment != null) continue;
            AddStrut(segment.Text, run.Style);
            foreach (HtmlRenderBoxStyle style in run.InlineStrutStyles) AddStrut(string.Empty, style);
        }
        foreach (HtmlRenderBoxStyle style in line.EmptyInlineStrutStyles) AddStrut(string.Empty, style);

        lineHeight = Math.Max(lineHeight, ascent + descent);
        // Retain the existing paragraph paint anchor. Only the extra top extent
        // lifts that anchor; the selected face supplies line advance, not a global
        // migration of ordinary text to browser-relative baseline coordinates.
        baseline = ResolveInlineTextBaseline(strutText, paragraphStyle, hasImage: false)
            + ascent - strutAscent;
        return true;

        void AddStrut(string text, HtmlRenderBoxStyle style) {
            double sourceAscent = ResolveTextLineAscent(text, style);
            double displacement = style.BaselineOffset - paragraphStyle.BaselineOffset;
            ascent = Math.Max(ascent, sourceAscent - displacement);
            descent = Math.Max(descent, style.LineHeight - sourceAscent + displacement);
        }
    }

    private double ResolveTextLineAscent(string text, HtmlRenderBoxStyle style) {
        HtmlTextFaceMetrics? face = ResolveTextFaceMetrics(text, style);
        return IsUsableTextFaceMetrics(face)
            ? (style.LineHeight - face!.Value.Height) / 2D + face.Value.BaselineOffset
            : ResolveTextAscent(style);
    }
}
