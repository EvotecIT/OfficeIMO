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
        bool isFinalContentToken,
        InlineWordFitting wordFitting,
        int sourceStartOffset,
        int sourceEndOffset,
        bool previewContinuation) {
        if (!token.HasBreaks || token.PaintText.Length != token.LogicalText.Length) return false;
        if (run.Style.HyphenateLimitLast == "always"
            && isFinalContentToken
            && line.HasFlowContent
            && wordFitting.CanBreakBefore(run, sourceStartOffset)
            && InlineLine.PreviewEmptyLineAdvance(run, MeasureInlineText(token.PaintText, run.Style)) <= line.AvailableWidth + 0.0001D) {
            CommitFloatLine(lines, ref line, ref y, context, lineHeight);
        }
        if (line.HasFlowContent && run.Style.HyphenateLimitZone > 0D
            && wordFitting.CanBreakBefore(run, sourceStartOffset)
            && line.AvailableWidth - line.Width <= run.Style.HyphenateLimitZone + 0.0001D) {
            CommitFloatLine(lines, ref line, ref y, context, lineHeight);
        }

        int start = 0;
        while (start < token.PaintText.Length) {
            double available = Math.Max(0D, line.AvailableWidth - line.PreviewAdvance(run, 0D, false));
            bool hyphenationAllowed = !run.Style.HyphenateLimitLines.HasValue
                || CountConsecutiveHyphenatedLines(lines) < run.Style.HyphenateLimitLines.Value;
            int selectedEnd = -1;
            bool selectedIsBreak = false;
            double remainderWidth = MeasureInlineText(token.PaintText.Substring(start), run.Style);
            double remainderAdvance = previewContinuation && start == 0
                ? wordFitting.PreviewWithContinuation(line, run, sourceEndOffset, remainderWidth)
                : line.PreviewAdvance(run, remainderWidth);
            if (remainderAdvance <= line.AvailableWidth + 0.0001D) {
                selectedEnd = token.PaintText.Length;
            } else if (hyphenationAllowed) {
                selectedEnd = SelectHyphenationBreak(token.PrimaryBreaks, token.PaintText, start, available, run.Style);
                if (selectedEnd < 0) selectedEnd = SelectHyphenationBreak(token.SecondaryBreaks, token.PaintText, start, available, run.Style);
                selectedIsBreak = selectedEnd >= 0;
            }

            if (selectedEnd < 0) {
                if (start == 0 && line.HasFlowContent && !wordFitting.CanBreakBefore(run, sourceStartOffset)) return false;
                if (line.HasFlowContent) {
                    CommitFloatLine(lines, ref line, ref y, context, lineHeight);
                    continue;
                }
                double remainingWidth = remainderWidth;
                double previousY = y;
                MoveFloatLineBelowObstruction(ref line, ref y, context, lineHeight, line.PreviewAdvance(run, remainingWidth));
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
        string logicalToken,
        bool followsCollapsibleSpace,
        InlineWordFitting wordFitting,
        int sourceStartOffset,
        int sourceEndOffset,
        IReadOnlyList<int> sourceBoundaries) {
        if (paintToken.Length != logicalToken.Length) return false;
        IReadOnlyList<int> breaks = GetHtmlPreferredBreakPositions(
            paintToken,
            allowCjkBreaks: run.Style.WordBreak != "keep-all");
        if (breaks.Count == 0) return false;

        int start = 0;
        foreach (int end in breaks.Concat(new[] { paintToken.Length })) {
            if (end <= start || end > paintToken.Length) continue;
            string paintChunk = paintToken.Substring(start, end - start);
            string logicalChunk = logicalToken.Substring(start, end - start);
            double chunkWidth = MeasureInlineText(paintChunk, run.Style);
            bool completesToken = end == paintToken.Length;
            int sourceStart = sourceStartOffset + (start < sourceBoundaries.Count ? sourceBoundaries[start] : start);
            int sourceEnd = completesToken ? sourceEndOffset
                : sourceStartOffset + (end < sourceBoundaries.Count ? sourceBoundaries[end] : end);
            bool canBreakBefore = wordFitting.CanBreakBefore(run, sourceStart);
            double chunkAdvance = canBreakBefore
                ? wordFitting.PreviewWithContinuation(line, run, sourceEnd, chunkWidth, completesToken)
                : line.PreviewAdvance(run, chunkWidth, completesToken);
            if (canBreakBefore && !line.HasFlowContent && chunkAdvance > line.AvailableWidth + 0.0001D) {
                MoveFloatLineBelowObstruction(ref line, ref y, context, lineHeight, chunkAdvance);
            }
            bool emergencyContinuation = !canBreakBefore && line.HasFlowContent
                && line.PreviewAdvance(run, chunkWidth, completesToken) > line.AvailableWidth + 0.0001D;
            if ((InlineLine.PreviewEmptyLineAdvance(run, chunkWidth, completesToken) > line.AvailableWidth + 0.0001D
                    || emergencyContinuation) && AllowsEmergencyTokenBreak(run.Style)) {
                if (start == 0 && followsCollapsibleSpace && line.HasFlowContent && run.Style.WordBreak != "break-all") {
                    CommitFloatLine(lines, ref line, ref y, context, lineHeight);
                }
                AddBrokenFloatToken(lines, ref line, ref y, context, lineHeight, run, paintChunk, completesToken);
                start = end;
                continue;
            }
            if (canBreakBefore && line.HasFlowContent && chunkAdvance > line.AvailableWidth + 0.0001D) {
                CommitFloatLine(lines, ref line, ref y, context, lineHeight);
            }
            line.Add(new InlineSegment(paintChunk, chunkWidth, run, logicalChunk));
            start = end;
        }
        return start == paintToken.Length;
    }
}
