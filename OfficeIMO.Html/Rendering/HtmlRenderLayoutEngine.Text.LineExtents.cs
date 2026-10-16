namespace OfficeIMO.Html {

    internal sealed partial class HtmlRenderLayoutEngine {
        /// <summary>
        /// Combines the ascent and descent of text-only horizontal lines before paint
        /// or float placement. Font size alone is not a face's ascent, and inherited
        /// leading belongs equally above and below each run's font box.
        /// </summary>
        private bool TryResolveTextOnlyLineMetrics(InlineLine line, HtmlRenderBoxStyle paragraphStyle,
            ref double lineHeight, out double baseline, out bool ordinaryLineCarrier) {
            baseline = 0D;
            ordinaryLineCarrier = false;
            if (!line.HasFlowContent || IsVerticalWritingMode(paragraphStyle.WritingMode)
                || line.Segments.Any(segment => segment.Run.AtomicBlock != null && !segment.Run.IsFlowMarker)) return false;
            string strutText = line.Segments.FirstOrDefault(segment => segment.Text.Length > 0)?.Text ?? string.Empty;
            bool ordinaryFontBoxes = CanUseOrdinaryInlineFontBoxes(line, paragraphStyle);
            bool selectedUsedFace = paragraphStyle.Font.Size > 0D
                && !line.Segments.Any(segment => segment.Run.LeaderPattern != null)
                && TryResolveNormalInlineTextMetrics(paragraphStyle, ResolveTextFaceMetrics(strutText, paragraphStyle), out _,
                    allowShortLine: ordinaryFontBoxes);
            if (!selectedUsedFace && !line.HasMixedTextSizes(paragraphStyle)
                    && !line.Segments.Any(segment => InlineLine.ContributesFlowContent(segment.Run)
                        && Math.Abs(segment.Run.Style.BaselineOffset
                        - paragraphStyle.BaselineOffset) > 0.000001D)) return false;
            ordinaryLineCarrier = selectedUsedFace;
            InlineTextLineMetrics strut = ResolveTextLineMetrics(strutText, paragraphStyle, ordinaryFontBoxes);
            double strutAscent = strut.Baseline;
            double ascent = strutAscent;
            double descent = (selectedUsedFace ? strut.Height : paragraphStyle.LineHeight) - strutAscent;
            double shortLineAscent = paragraphStyle.LineHeight < paragraphStyle.Font.Size
                ? strutAscent : double.NegativeInfinity;
            foreach (InlineSegment segment in line.Segments) {
                HtmlInlineRun run = segment.Run;
                if (!InlineLine.ContributesFlowContent(run)) continue;
                AddStrut(segment.Text, run.Style);
                foreach (HtmlRenderBoxStyle style in run.InlineStrutStyles) AddStrut(string.Empty, style);
            }
            foreach (HtmlRenderBoxStyle style in line.EmptyInlineStrutStyles) AddStrut(string.Empty, style);

            // The normal selected-face route owns a used line height already.
            // Its raw computed CSS length must not remain a second flow minimum.
            lineHeight = selectedUsedFace ? ascent + descent : Math.Max(lineHeight, ascent + descent);
            // A selected face supplies both its normal baseline and the line's
            // participating extents. Unavailable faces retain their existing
            // paragraph anchor; specialized short lines retain their raw-face route.
            baseline = ordinaryFontBoxes || Math.Abs(shortLineAscent - ascent) < 0.000001D ? ascent
                : ResolveInlineTextBaseline(strutText, paragraphStyle, hasImage: false) + ascent - strutAscent;
            return true;

            void AddStrut(string text, HtmlRenderBoxStyle style) {
                InlineTextLineMetrics source = ResolveTextLineMetrics(text, style, ordinaryFontBoxes);
                double sourceAscent = source.Baseline;
                double displacement = style.BaselineOffset - paragraphStyle.BaselineOffset;
                if (style.LineHeight < style.Font.Size) shortLineAscent = Math.Max(shortLineAscent, sourceAscent - displacement);
                ascent = Math.Max(ascent, sourceAscent - displacement);
                descent = Math.Max(descent, (selectedUsedFace ? source.Height : style.LineHeight) - sourceAscent + displacement);
            }
        }

        private InlineTextLineMetrics ResolveTextLineMetrics(string text, HtmlRenderBoxStyle style, bool allowShortLine = false) {
            // Zero-size ancestors still contribute inherited leading around visible
            // descendants, but cannot be passed to a positive-size font resolver.
            if (style.Font.Size == 0D) return new InlineTextLineMetrics(style.LineHeight, style.LineHeight / 2D);
            HtmlTextFaceMetrics? face = ResolveTextFaceMetrics(text, style);
            if (TryResolveNormalInlineTextMetrics(style, face, out InlineTextLineMetrics normal, allowShortLine)) return normal;
            double baseline = IsUsableTextFaceMetrics(face)
                ? (style.LineHeight - face!.Value.Height) / 2D + face.Value.BaselineOffset
                : style.LineHeight / 2D + GetEffectiveTextFont(style).Size * 0.3D;
            return new InlineTextLineMetrics(style.LineHeight, baseline);
        }
    }
}
