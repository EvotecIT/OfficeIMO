using System.Globalization;
using System.Text;
using AngleSharp.Dom;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private bool TryAddHyphenatedToken(
        ICollection<InlineLine> lines,
        ref InlineLine line,
        HtmlInlineRun run,
        HyphenationToken token,
        double width,
        int logicalStartProgress,
        int logicalEndProgress,
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
            && InlineLine.PreviewEmptyLineAdvance(run, MeasureInlineText(token.PaintText, run.Style)) <= width + 0.0001D) {
            TrimTrailingWhitespace(line);
            lines.Add(line);
            line = new InlineLine();
            line.Add(new InlineSegment(
                token.PaintText,
                MeasureInlineText(token.PaintText, run.Style),
                run,
                token.LogicalText,
                logicalEndProgress: logicalEndProgress));
            return true;
        }
        if (line.HasFlowContent && run.Style.HyphenateLimitZone > 0D
            && wordFitting.CanBreakBefore(run, sourceStartOffset)
            && line.ResolveAvailableWidth(width) - line.Width <= run.Style.HyphenateLimitZone + 0.0001D) {
            TrimTrailingWhitespace(line);
            lines.Add(line);
            line = new InlineLine();
        }

        int start = 0;
        while (start < token.PaintText.Length) {
            double available = Math.Max(0D, line.ResolveAvailableWidth(width) - line.PreviewAdvance(run, 0D, false));
            bool hyphenationAllowed = !run.Style.HyphenateLimitLines.HasValue
                || CountConsecutiveHyphenatedLines(lines) < run.Style.HyphenateLimitLines.Value;
            int selectedEnd = -1;
            bool selectedIsBreak = false;
            double remainderWidth = MeasureInlineText(token.PaintText.Substring(start), run.Style);
            double remainderAdvance = previewContinuation && start == 0
                ? wordFitting.PreviewWithContinuation(line, run, sourceEndOffset, remainderWidth)
                : line.PreviewAdvance(run, remainderWidth);
            if (remainderAdvance <= line.ResolveAvailableWidth(width) + 0.0001D) {
                selectedEnd = token.PaintText.Length;
            } else if (hyphenationAllowed) {
                selectedEnd = SelectHyphenationBreak(token.PrimaryBreaks, token.PaintText, start, available, run.Style);
                if (selectedEnd < 0) {
                    selectedEnd = SelectHyphenationBreak(token.SecondaryBreaks, token.PaintText, start, available, run.Style);
                }
                selectedIsBreak = selectedEnd >= 0;
            }

            if (selectedEnd < 0) {
                if (start == 0 && line.HasFlowContent && !wordFitting.CanBreakBefore(run, sourceStartOffset)) return false;
                if (!line.HasFlowContent) {
                    if (AllowsEmergencyTokenBreak(run.Style)) return false;
                    string paintRemainder = token.PaintText.Substring(start);
                    string logicalRemainder = token.LogicalText.Substring(start);
                    line.Add(new InlineSegment(
                        paintRemainder,
                        MeasureInlineText(paintRemainder, run.Style),
                        run,
                        logicalRemainder,
                        logicalEndProgress: logicalEndProgress));
                    return true;
                }
                TrimTrailingWhitespace(line);
                lines.Add(line);
                line = new InlineLine();
                continue;
            }

            string paintChunk = token.PaintText.Substring(start, selectedEnd - start)
                + (selectedIsBreak ? run.Style.HyphenateCharacter : string.Empty);
            string logicalChunk = token.LogicalText.Substring(start, selectedEnd - start);
            int sourceBoundary = selectedEnd < token.SourceBoundaries.Count
                ? token.SourceBoundaries[selectedEnd]
                : logicalEndProgress - logicalStartProgress;
            line.Add(new InlineSegment(
                paintChunk,
                MeasureInlineText(paintChunk, run.Style),
                run,
                logicalChunk,
                logicalEndProgress: selectedEnd == token.PaintText.Length
                    ? logicalEndProgress
                    : logicalStartProgress + sourceBoundary));
            start = selectedEnd;
            if (selectedIsBreak) {
                line.EndsWithHyphenation = true;
                lines.Add(line);
                line = new InlineLine();
            }
        }
        return true;
    }

    private int SelectHyphenationBreak(
        IReadOnlyList<int> candidates,
        string paintText,
        int start,
        double available,
        HtmlRenderBoxStyle style) {
        int selected = -1;
        foreach (int point in candidates) {
            if (point <= start) continue;
            string candidate = paintText.Substring(start, point - start) + style.HyphenateCharacter;
            if (MeasureInlineText(candidate, style) <= available + 0.0001D) selected = point;
        }
        return selected;
    }

    private static int CountConsecutiveHyphenatedLines(ICollection<InlineLine> lines) {
        int count = 0;
        foreach (InlineLine candidate in lines.Reverse()) {
            if (!candidate.EndsWithHyphenation) break;
            count++;
        }
        return count;
    }

    private bool TryAddPreferredBreakToken(
        ICollection<InlineLine> lines,
        ref InlineLine line,
        HtmlInlineRun run,
        string paintToken,
        string logicalToken,
        double width,
        int logicalStartProgress,
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
            bool emergencyContinuation = !canBreakBefore && line.HasFlowContent
                && line.PreviewAdvance(run, chunkWidth, completesToken) > line.ResolveAvailableWidth(width) + 0.0001D;
            if ((InlineLine.PreviewEmptyLineAdvance(run, chunkWidth, completesToken) > line.ResolveAvailableWidth(width) + 0.0001D
                    || emergencyContinuation) && AllowsEmergencyTokenBreak(run.Style)) {
                if (start == 0 && followsCollapsibleSpace && line.HasFlowContent && run.Style.WordBreak != "break-all") {
                    TrimTrailingWhitespace(line);
                    lines.Add(line);
                    line = new InlineLine();
                }
                AddBrokenToken(lines, ref line, run, paintChunk, logicalChunk, width, logicalStartProgress + start, completesToken);
                start = end;
                continue;
            }
            if (canBreakBefore && line.HasFlowContent && chunkAdvance > line.ResolveAvailableWidth(width) + 0.0001D) {
                TrimTrailingWhitespace(line);
                lines.Add(line);
                line = new InlineLine();
            }
            line.Add(new InlineSegment(
                paintChunk,
                chunkWidth,
                run,
                logicalChunk,
                logicalEndProgress: logicalStartProgress + (end < sourceBoundaries.Count ? sourceBoundaries[end] : end)));
            start = end;
        }
        return start == paintToken.Length;
    }

    private IReadOnlyList<int> GetHtmlPreferredBreakPositions(string text, bool allowCjkBreaks) {
        // Normal HTML line layout keeps unspaced solidus text together in Chromium.
        // Explicit overflow-wrap and word-break modes can still split it later.
        return OfficeTextLineBreaks.GetBreakPositions(text, allowCjkBreaks, _cancellationToken)
            .Where(point => text[point - 1] != '/')
            .ToArray();
    }

    private double MeasureTabExpandedText(string value, HtmlRenderBoxStyle style, double currentWidth) {
        if (value.IndexOf('\t') < 0) return MeasureInlineText(value, style);
        double spaceWidth = Math.Max(0.01D, MeasureInlineText(" ", style));
        double stopWidth = style.TabSizeIsLength ? style.TabSize : style.TabSize * spaceWidth;
        double cursor = currentWidth;
        foreach (char character in value) {
            if (character != '\t') {
                cursor += MeasureInlineText(character.ToString(), style);
                continue;
            }
            if (stopWidth <= 0D) continue;
            double nextStop = (Math.Floor(cursor / stopWidth) + 1D) * stopWidth;
            cursor = nextStop;
        }
        return Math.Max(0D, cursor - currentWidth);
    }

    private void AddTabExpandedSegments(
        InlineLine line,
        string paintText,
        string logicalText,
        HtmlInlineRun run,
        int logicalStartProgress = 0) {
        double spaceWidth = Math.Max(0.01D, MeasureInlineText(" ", run.Style));
        double stopWidth = run.Style.TabSizeIsLength ? run.Style.TabSize : run.Style.TabSize * spaceWidth;
        int chunkStart = 0;
        int logicalProgress = logicalStartProgress;
        for (int index = 0; index < paintText.Length; index++) {
            if (paintText[index] != '\t') continue;
            if (index > chunkStart) {
                string paintChunk = paintText.Substring(chunkStart, index - chunkStart);
                string logicalChunk = logicalText.Substring(chunkStart, index - chunkStart);
                logicalProgress += logicalChunk.Length;
                line.Add(new InlineSegment(
                    paintChunk,
                    MeasureInlineText(paintChunk, run.Style),
                    run,
                    logicalChunk,
                    logicalEndProgress: logicalProgress));
            }

            double tabWidth = 0D;
            if (stopWidth > 0D) {
                double textStart = line.PreviewContentStart(run);
                double nextStop = (Math.Floor(textStart / stopWidth) + 1D) * stopWidth;
                tabWidth = Math.Max(0D, nextStop - textStart);
            }
            string logicalTab = logicalText.Substring(index, 1);
            logicalProgress += logicalTab.Length;
            line.Add(new InlineSegment(string.Empty, tabWidth, run, logicalTab, logicalEndProgress: logicalProgress));
            chunkStart = index + 1;
        }

        if (chunkStart >= paintText.Length) return;
        string finalPaint = paintText.Substring(chunkStart);
        string finalLogical = logicalText.Substring(chunkStart);
        logicalProgress += finalLogical.Length;
        line.Add(new InlineSegment(
            finalPaint,
            MeasureInlineText(finalPaint, run.Style),
            run,
            finalLogical,
            logicalEndProgress: logicalProgress));
    }

    private void ApplyEndEllipsis(InlineLine line, double width, int completeLogicalProgress) {
        double availableWidth = line.ResolveAvailableWidth(width);
        TrimTrailingWhitespace(line);
        HtmlInlineRun? ellipsisRun = line.Segments
            .LastOrDefault(segment => segment.Run.AtomicBlock == null && segment.Text.Length > 0)?.Run;
        if (ellipsisRun == null && line.Segments.Count > 0) ellipsisRun = line.Segments[line.Segments.Count - 1].Run;
        while (line.Segments.Count > 0) {
            InlineSegment segment = line.Segments[line.Segments.Count - 1];
            if (segment.Run.AtomicBlock != null || segment.Text.Length == 0) {
                line.RemoveAt(line.Segments.Count - 1);
                continue;
            }

            ellipsisRun = segment.Run;
            line.RemoveAt(line.Segments.Count - 1);
            double remainingWidth = Math.Max(0D, availableWidth - line.Width);
            double ellipsisWidth = MeasureInlineText("\u2026", segment.Run.Style);
            if (ellipsisWidth > remainingWidth + 0.0001D) continue;

            var paint = new StringBuilder();
            var logical = new StringBuilder();
            IReadOnlyList<string> paintElements = OfficeTextElements.Split(segment.Text);
            IReadOnlyList<string> logicalElements = OfficeTextElements.Split(segment.LogicalText);
            for (int index = 0; index < paintElements.Count; index++) {
                string candidate = paint.ToString() + paintElements[index];
                if (MeasureInlineText(candidate, segment.Run.Style) + ellipsisWidth > remainingWidth + 0.0001D) break;
                paint.Append(paintElements[index]);
                if (index < logicalElements.Count) logical.Append(logicalElements[index]);
            }

            string text = paint.ToString() + "\u2026";
            string logicalText = logical.ToString() + "\u2026";
            line.Add(new InlineSegment(
                text,
                MeasureInlineText(text, segment.Run.Style),
                segment.Run,
                logicalText,
                logicalEndProgress: completeLogicalProgress));
            return;
        }

        if (ellipsisRun != null) {
            ellipsisRun = CreateEllipsisRun(ellipsisRun);
            double ellipsisWidth = MeasureInlineText("\u2026", ellipsisRun.Style);
            if (ellipsisWidth <= availableWidth + 0.0001D) {
                line.Add(new InlineSegment("\u2026", ellipsisWidth, ellipsisRun, "\u2026", logicalEndProgress: completeLogicalProgress));
            }
        }
    }

    private static HtmlInlineRun CreateEllipsisRun(HtmlInlineRun source) {
        if (source.AtomicBlock == null) return source;
        var run = new HtmlInlineRun(
            "\u2026",
            source.Style,
            source.LinkUri,
            source.Source,
            source.PaintOffsetX,
            source.PaintOffsetY,
            source.OwnerElement,
            logicalText: "\u2026");
        if (source.SemanticNodeId.HasValue) run.AssignSemanticNode(source.SemanticRole, source.SemanticNodeId.Value, source.BookmarkAnchorText, source.SemanticFragmentOrder);
        if (source.LogicalTextOrder.HasValue) run.AssignLogicalTextOrder(source.LogicalTextOrder.Value);
        if (source.InlineSemanticGroupRole.HasValue && source.InlineSemanticGroupKey != null) {
            run.AssignInlineSemanticGroup(source.InlineSemanticGroupRole.Value, source.InlineSemanticGroupKey);
        }
        return run;
    }

    private static bool AllowsEmergencyTokenBreak(HtmlRenderBoxStyle style) =>
        style.OverflowWrap == "anywhere"
        || style.OverflowWrap == "break-word"
        || style.WordBreak == "break-all"
        || style.WordBreak == "break-word";

    private static IReadOnlyList<InlineSegment> MergeAdjacentInlineSegments(IReadOnlyList<InlineSegment> segments) {
        var merged = new List<InlineSegment>(segments.Count);
        foreach (InlineSegment segment in segments) {
            if (segment.Run.AtomicBlock == null
                && segment.Text.Length > 0
                && merged.Count > 0
                && merged[merged.Count - 1].Text.Length > 0
                && ReferenceEquals(merged[merged.Count - 1].Run, segment.Run)) {
                InlineSegment previous = merged[merged.Count - 1];
                merged[merged.Count - 1] = new InlineSegment(
                    previous.Text + segment.Text,
                    previous.Width + segment.Width,
                    previous.Run,
                    previous.LogicalText + segment.LogicalText,
                    logicalEndProgress: segment.LogicalEndProgress) {
                    LeadingAdvance = previous.LeadingAdvance + segment.LeadingAdvance,
                    TrailingAdvance = previous.TrailingAdvance + segment.TrailingAdvance
                };
            } else {
                merged.Add(segment);
            }
        }

        return merged;
    }

    private void AddBrokenToken(
        ICollection<InlineLine> lines,
        ref InlineLine line,
        HtmlInlineRun run,
        string token,
        string logicalToken,
        double width,
        int logicalStartProgress,
        bool completesToken = true) {
        var part = new StringBuilder();
        var logicalPart = new StringBuilder();
        double partWidth = 0D;
        int partLogicalLength = 0;
        IReadOnlyList<string> paintElements = OfficeTextElements.Split(token);
        IReadOnlyList<string> logicalElements = OfficeTextElements.Split(logicalToken);
        for (int index = 0; index < paintElements.Count; index++) {
            string value = paintElements[index];
            string logicalValue = index < logicalElements.Count ? logicalElements[index] : OfficeArabicTextShaper.ToLogicalText(value);
            double charWidth = MeasureInlineText(value, run.Style);
            bool completesPiece = completesToken && index == paintElements.Count - 1;
            if (part.Length > 0 && line.PreviewAdvance(run, partWidth + charWidth, completesPiece) > line.ResolveAvailableWidth(width) + 0.0001D) {
                line.Add(new InlineSegment(
                    part.ToString(),
                    partWidth,
                    run,
                    logicalPart.ToString(),
                    logicalEndProgress: logicalStartProgress + partLogicalLength));
                lines.Add(line);
                line = new InlineLine();
                part.Clear();
                logicalPart.Clear();
                partWidth = 0D;
            } else if (part.Length == 0 && line.HasFlowContent && line.PreviewAdvance(run, charWidth, completesPiece) > line.ResolveAvailableWidth(width) + 0.0001D) {
                TrimTrailingWhitespace(line);
                lines.Add(line);
                line = new InlineLine();
            }

            part.Append(value);
            logicalPart.Append(logicalValue);
            partLogicalLength += logicalValue.Length;
            partWidth += charWidth;
        }

        if (part.Length > 0) {
            if (line.HasFlowContent && line.PreviewAdvance(run, partWidth, completesToken) > line.ResolveAvailableWidth(width) + 0.0001D) {
                TrimTrailingWhitespace(line);
                lines.Add(line);
                line = new InlineLine();
            }

            line.Add(new InlineSegment(
                part.ToString(),
                partWidth,
                run,
                logicalPart.ToString(),
                logicalEndProgress: logicalStartProgress + partLogicalLength));
        }
    }

    private static string SliceLogicalToken(HtmlInlineRun run, string token, ref int offset) {
        if (offset >= 0 && token.Length <= run.LogicalText.Length - offset) {
            string value = run.LogicalText.Substring(offset, token.Length);
            offset += token.Length;
            return value;
        }

        offset += token.Length;
        return OfficeArabicTextShaper.ToLogicalText(token);
    }

}
