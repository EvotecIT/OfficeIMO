using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private bool TryAddHyphenatedFloatToken(
        ICollection<InlineLine> lines,
        ref InlineLine line,
        ref double y,
        InlineFloatContext context,
        double lineHeight,
        HtmlInlineRun run,
        HyphenationToken token,
        bool isFinalContentToken) {
        if (!token.HasBreaks || token.PaintText.Length != token.LogicalText.Length) return false;
        if (run.Style.HyphenateLimitLast == "always"
            && isFinalContentToken
            && line.HasFlowContent
            && MeasureInlineText(token.PaintText, run.Style) <= line.AvailableWidth + 0.0001D) {
            CommitFloatLine(lines, ref line, ref y, context, lineHeight);
        }
        if (line.HasFlowContent && run.Style.HyphenateLimitZone > 0D
            && line.AvailableWidth - line.Width <= run.Style.HyphenateLimitZone + 0.0001D) {
            CommitFloatLine(lines, ref line, ref y, context, lineHeight);
        }

        int start = 0;
        while (start < token.PaintText.Length) {
            double available = Math.Max(0D, line.AvailableWidth - line.Width);
            bool hyphenationAllowed = !run.Style.HyphenateLimitLines.HasValue
                || CountConsecutiveHyphenatedLines(lines) < run.Style.HyphenateLimitLines.Value;
            int selectedEnd = -1;
            bool selectedIsBreak = false;
            if (MeasureInlineText(token.PaintText.Substring(start), run.Style) <= available + 0.0001D) {
                selectedEnd = token.PaintText.Length;
            } else if (hyphenationAllowed) {
                selectedEnd = SelectHyphenationBreak(token.PrimaryBreaks, token.PaintText, start, available, run.Style);
                if (selectedEnd < 0) selectedEnd = SelectHyphenationBreak(token.SecondaryBreaks, token.PaintText, start, available, run.Style);
                selectedIsBreak = selectedEnd >= 0;
            }

            if (selectedEnd < 0) {
                if (line.HasFlowContent) {
                    CommitFloatLine(lines, ref line, ref y, context, lineHeight);
                    continue;
                }
                double remainingWidth = MeasureInlineText(token.PaintText.Substring(start), run.Style);
                double previousY = y;
                MoveFloatLineBelowObstruction(ref line, ref y, context, lineHeight, remainingWidth);
                if (y > previousY + 0.0001D) continue;
                if (AllowsEmergencyTokenBreak(run.Style)) return false;
                string paintRemainder = token.PaintText.Substring(start);
                string logicalRemainder = token.LogicalText.Substring(start);
                line.Add(new InlineSegment(paintRemainder, remainingWidth, run, logicalRemainder));
                return true;
            }

            string paintChunk = token.PaintText.Substring(start, selectedEnd - start)
                + (selectedIsBreak ? run.Style.HyphenateCharacter : string.Empty);
            string logicalChunk = token.LogicalText.Substring(start, selectedEnd - start);
            line.Add(new InlineSegment(paintChunk, MeasureInlineText(paintChunk, run.Style), run, logicalChunk));
            start = selectedEnd;
            if (selectedIsBreak) {
                line.EndsWithHyphenation = true;
                CommitFloatLine(lines, ref line, ref y, context, lineHeight);
            }
        }
        return true;
    }

    private bool TryAddPreferredFloatBreakToken(
        ICollection<InlineLine> lines,
        ref InlineLine line,
        ref double y,
        InlineFloatContext context,
        double lineHeight,
        HtmlInlineRun run,
        string paintToken,
        string logicalToken) {
        if (paintToken.Length != logicalToken.Length) return false;
        IReadOnlyList<int> breaks = OfficeTextLineBreaks.GetBreakPositions(
            paintToken,
            allowCjkBreaks: run.Style.WordBreak != "keep-all");
        if (breaks.Count == 0) return false;

        int start = 0;
        foreach (int end in breaks.Concat(new[] { paintToken.Length })) {
            if (end <= start || end > paintToken.Length) continue;
            string paintChunk = paintToken.Substring(start, end - start);
            string logicalChunk = logicalToken.Substring(start, end - start);
            double chunkWidth = MeasureInlineText(paintChunk, run.Style);
            if (!line.HasFlowContent && chunkWidth > line.AvailableWidth + 0.0001D) {
                MoveFloatLineBelowObstruction(ref line, ref y, context, lineHeight, chunkWidth);
            }
            if (chunkWidth > line.AvailableWidth + 0.0001D && AllowsEmergencyTokenBreak(run.Style)) {
                AddBrokenFloatToken(lines, ref line, ref y, context, lineHeight, run, paintChunk);
                start = end;
                continue;
            }
            if (line.HasFlowContent && line.Width + chunkWidth > line.AvailableWidth + 0.0001D) {
                CommitFloatLine(lines, ref line, ref y, context, lineHeight);
            }
            line.Add(new InlineSegment(paintChunk, chunkWidth, run, logicalChunk));
            start = end;
        }
        return start == paintToken.Length;
    }
}
