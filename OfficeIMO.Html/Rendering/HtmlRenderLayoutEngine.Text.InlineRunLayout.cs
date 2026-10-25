using System.Globalization;
using System.Text;
using AngleSharp.Dom;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private HtmlInlineLayout LayoutInlineRuns(
        IReadOnlyList<HtmlInlineRun> runs,
        double width,
        HtmlRenderBoxStyle paragraphStyle,
        IElement? formattingContainer = null,
        int skipLogicalCharacters = 0,
        PagedFloatBoundary? pageBoundary = null,
        IReadOnlyList<HtmlFloatExclusion>? inheritedFloats = null,
        InlinePaintCapture? paintCapture = null, InlineFloatContext? floatContext = null,
        bool inlineEdgesPrepared = false, int? maximumLines = null, Action? onLineLimitExceeded = null,
        bool retainPreparedTokenClusters = false) {
        AssignLogicalTextOrders(runs);
        if (!inlineEdgesPrepared) PrepareInlineEdgeScopes(runs);
        if (runs.Count == 0 || width <= 0D) return new HtmlInlineLayout(Array.Empty<HtmlRenderVisual>(), 0D,
            hasInFlowLineBoxes: runs.Count == 0 ? false : (bool?)null);
        if (runs.Any(run => run.IsBlockInterruption)) {
            return LayoutInterruptedInlineRuns(runs, width, paragraphStyle, formattingContainer);
        }
        if (floatContext?.HasFloats == true || runs.Any(run => run.FloatingBlock != null) || inheritedFloats?.Count > 0) {
            return LayoutInlineRunsWithFloats(runs, width, paragraphStyle, formattingContainer, pageBoundary, inheritedFloats, paintCapture, floatContext);
        }
        bool supportsContinuationReflow = runs.All(run =>
            run.AtomicBlock == null
            && run.PositionedMarkerElement == null
            && run.RunningStringElement == null
            && run.RunningElementAssignment == null
            && run.Text.IndexOf('\u2028') < 0
            && run.Text.IndexOf('\n') < 0
            && run.Text.IndexOf('\r') < 0);
        int canonicalProgress = 0;
        bool canonicalHasContent = false;
        bool canonicalPreviousWasCollapsibleSpace = false;
        var lines = new List<InlineLine>();
        ICollection<InlineLine> lineAdmission = maximumLines.HasValue
            ? new InlineLineCollection(lines, maximumLines.Value, retainPreparedTokenClusters) : lines;
        var line = new InlineLine();
        if (skipLogicalCharacters == 0) line.Indent(paragraphStyle.TextIndent?.Resolve(width) ?? 0D, paragraphStyle.Direction == "rtl");
        bool previousWasCollapsibleSpace = false;
        int noWrapRangeStart = -1;
        int noWrapStrutStart = 0;
        bool noWrapRangeStartedAfterContent = false;
        InlineWordFitting wordFitting = CreateInlineWordFitting(runs, paragraphStyle,
            consecutiveHyphenatedLines: () => CountConsecutiveHyphenatedLines(lines));
        try {
        for (int runIndex = 0; runIndex < runs.Count; runIndex++) {
            HtmlInlineRun run = runs[runIndex];
            if (ProcessInlineEdgeBoundary(run, line)) continue;
            bool runPreventsWrapping = !paragraphStyle.PreventTextWrapping && run.Style.PreventTextWrapping;
            if (!runPreventsWrapping && noWrapRangeStart >= 0) {
                previousWasCollapsibleSpace = FinalizeNoWrapRange(
                    lineAdmission,
                    ref line,
                    noWrapRangeStart,
                    noWrapStrutStart,
                    noWrapRangeStartedAfterContent,
                    width);
                noWrapRangeStart = -1;
                noWrapRangeStartedAfterContent = false;
            } else if (runPreventsWrapping && noWrapRangeStart < 0) {
                noWrapRangeStart = line.Segments.Count;
                noWrapStrutStart = line.EmptyInlineStrutStyles.Count;
                noWrapRangeStartedAfterContent = line.HasFlowContent;
            }
            if (run.IsInlineStrutMarker) {
                line.RecordEmptyInlineStruts(run.InlineStrutStyles);
                continue;
            }
            if (run.RunningStringElement != null) {
                line.Add(new InlineSegment(string.Empty, 0D, run));
                continue;
            }
            if (run.RunningElementAssignment != null) {
                line.Add(new InlineSegment(string.Empty, 0D, run));
                continue;
            }
            if (run.PositionedMarkerElement != null) {
                line.Add(new InlineSegment(string.Empty, 0D, run));
                previousWasCollapsibleSpace = false;
                continue;
            }
            if (run.LeaderPattern != null) {
                line.Add(new InlineSegment(string.Empty, 0D, run, string.Empty));
                continue;
            }
            if (run.AtomicBlock != null) {
                previousWasCollapsibleSpace = false;
                double atomicWidth = run.AtomicBlock.Width;
                if (!paragraphStyle.PreventTextWrapping
                    && !runPreventsWrapping
                    && line.HasFlowContent
                    && line.PreviewAdvance(run, atomicWidth) > line.ResolveAvailableWidth(width) + 0.0001D) {
                    TrimTrailingWhitespace(line);
                    lineAdmission.Add(line);
                    line = new InlineLine();
                }

                line.Add(new InlineSegment(string.Empty, atomicWidth, run));
                continue;
            }

            int logicalOffset = 0;
            int sourceOffset = 0;
            bool preserveWhitespace = run.Style.PreserveWhitespace;
            IReadOnlyList<string> tokens = Tokenize(run.Text, preserveWhitespace, run.Style.BreakSpaces).ToList();
            for (int tokenIndex = 0; tokenIndex < tokens.Count; tokenIndex++) {
                string token = tokens[tokenIndex];
                int sourceStartOffset = sourceOffset;
                sourceOffset += preserveWhitespace && token == "\n" && run.Text[sourceOffset] == '\r'
                    && sourceOffset + 1 < run.Text.Length && run.Text[sourceOffset + 1] == '\n' ? 2 : token.Length;
                run.InlineTokenEndsRun = tokenIndex == tokens.Count - 1;
                string logicalToken = SliceLogicalToken(run, token, ref logicalOffset);
                if (token == "\u2028" || preserveWhitespace && (token == "\n" || token == "\r\n")) {
                    line.HasForcedBreak = true;
                    if (noWrapRangeStart >= 0) {
                        FinalizeNoWrapRange(
                            lineAdmission,
                            ref line,
                            noWrapRangeStart,
                            noWrapStrutStart,
                            noWrapRangeStartedAfterContent,
                            width);
                    }
                    // The break belongs to the committed line even when its inline
                    // has no preceding glyphs. Carry its face and ancestor struts
                    // without painting text or advancing the content cursor.
                    line.Add(new InlineSegment(string.Empty, 0D, run, string.Empty));
                    lineAdmission.Add(line);
                    line = new InlineLine();
                    previousWasCollapsibleSpace = false;
                    noWrapRangeStart = runPreventsWrapping ? 0 : -1;
                    noWrapStrutStart = 0;
                    noWrapRangeStartedAfterContent = false;
                    continue;
                }

                bool whitespace = IsWhitespaceToken(token);
                string normalizedToken = !preserveWhitespace && whitespace ? " " : token;
                string normalizedLogicalToken = !preserveWhitespace && whitespace ? " " : logicalToken;
                bool followsCollapsibleSpace = !whitespace && !preserveWhitespace && previousWasCollapsibleSpace;
                bool contributesCanonicalProgress = preserveWhitespace
                    || !whitespace
                    || canonicalHasContent && !canonicalPreviousWasCollapsibleSpace;
                int tokenStart = canonicalProgress;
                if (contributesCanonicalProgress) canonicalProgress += normalizedLogicalToken.Length;
                int tokenEnd = canonicalProgress;
                if (!preserveWhitespace) {
                    if (whitespace) {
                        canonicalPreviousWasCollapsibleSpace = true;
                    } else {
                        canonicalHasContent = true;
                        canonicalPreviousWasCollapsibleSpace = false;
                    }
                }

                if (!preserveWhitespace && whitespace) {
                    if (!line.HasFlowContent || previousWasCollapsibleSpace) continue;
                    previousWasCollapsibleSpace = true;
                } else {
                    previousWasCollapsibleSpace = false;
                }

                int visibleTokenStart = tokenStart;
                if (skipLogicalCharacters > tokenStart) {
                    int skipWithinToken = skipLogicalCharacters - tokenStart;
                    if (skipWithinToken >= normalizedLogicalToken.Length) {
                        MarkSkippedInlineEdges(run);
                        continue;
                    }
                    normalizedToken = normalizedToken.Substring(skipWithinToken);
                    MarkSkippedInlineEdges(run);
                    normalizedLogicalToken = normalizedLogicalToken.Substring(skipWithinToken);
                    visibleTokenStart += skipWithinToken;
                    whitespace = IsWhitespaceToken(normalizedToken);
                }

                bool hasTabs = preserveWhitespace && normalizedToken.IndexOf('\t') >= 0;
                double tabExpandedWidth = hasTabs ? MeasureTabExpandedText(normalizedToken, run.Style, line.PreviewContentStart(run)) : 0D;
                string paintToken = hasTabs ? normalizedToken.Replace("\t", string.Empty) : normalizedToken;
                HyphenationToken hyphenation = run.PreparedHyphenation.HasValue
                    ? run.PreparedHyphenation.Value.SliceSource(
                        logicalOffset - token.Length + visibleTokenStart - tokenStart, normalizedLogicalToken.Length)
                    : PrepareHyphenationToken(paintToken, normalizedLogicalToken, run.Style);
                paintToken = hyphenation.PaintText;
                string logicalPaintToken = hyphenation.LogicalText;
                double measured = hasTabs ? tabExpandedWidth : MeasureInlineText(paintToken, run.Style, run);
                if (run.EndsFirstLine && tokenIndex == tokens.Count - 1) {
                    paintToken += run.FirstLineHyphen;
                    line.Add(new InlineSegment(paintToken, MeasureInlineText(paintToken, run.Style, run), run,
                        logicalPaintToken, logicalEndProgress: tokenEnd));
                    line.EndsWithHyphenation = run.FirstLineHyphen.Length > 0;
                    lineAdmission.Add(line);
                    line = new InlineLine();
                    continue;
                }
                bool preventTokenWrapping = paragraphStyle.PreventTextWrapping || runPreventsWrapping;
                int visibleSourceStart = sourceStartOffset + visibleTokenStart - tokenStart;
                bool canBreakBefore = whitespace || wordFitting.CanBreakBefore(run, visibleSourceStart);
                bool previewContinuation = !preventTokenWrapping && !whitespace
                    && !AllowsEmergencyTokenBreak(run.Style)
                    && (canBreakBefore || hyphenation.HasBreaks
                        || wordFitting.HasBreakInside(run, visibleSourceStart, sourceOffset));
                double previewAdvance = previewContinuation
                    ? wordFitting.PreviewWithContinuation(line, run, sourceOffset, measured)
                    : line.PreviewAdvance(run, measured);
                if (!preventTokenWrapping
                    && !whitespace
                    && run.Style.WordBreak != "break-all"
                    && previewAdvance > line.ResolveAvailableWidth(width) + 0.0001D
                    && TryAddHyphenatedToken(
                        lineAdmission,
                        ref line,
                        run,
                        hyphenation,
                        width,
                        visibleTokenStart,
                        tokenEnd,
                        !HasRemainingInlineFlowContent(runs, runIndex, tokens, tokenIndex),
                        wordFitting, visibleSourceStart, sourceOffset, previewContinuation)) {
                    continue;
                }
                if (!preventTokenWrapping
                    && !whitespace
                    && previewAdvance > line.ResolveAvailableWidth(width) + 0.0001D
                    && TryAddPreferredBreakToken(
                        lineAdmission,
                        ref line,
                        run,
                        paintToken,
                        logicalPaintToken,
                        width,
                        visibleTokenStart,
                        followsCollapsibleSpace,
                        wordFitting, visibleSourceStart, sourceOffset, hyphenation.SourceBoundaries)) {
                    continue;
                }
                bool breakAllIntoRemainingSpace = run.Style.WordBreak == "break-all"
                    && line.HasFlowContent
                    && measured > Math.Max(0D, line.ResolveAvailableWidth(width) - line.PreviewAdvance(run, 0D)) + 0.0001D;
                bool continuationNeedsEmergencyBreak = !canBreakBefore && line.HasFlowContent
                    && line.PreviewAdvance(run, measured) > line.ResolveAvailableWidth(width) + 0.0001D;
                if (!preventTokenWrapping
                    && !whitespace
                    && AllowsEmergencyTokenBreak(run.Style)
                    && (InlineLine.PreviewEmptyLineAdvance(run, measured) > line.ResolveAvailableWidth(width) + 0.0001D
                        || breakAllIntoRemainingSpace || continuationNeedsEmergencyBreak)) {
                    if (followsCollapsibleSpace && line.HasFlowContent && run.Style.WordBreak != "break-all") {
                        TrimTrailingWhitespace(line);
                        lineAdmission.Add(line);
                        line = new InlineLine();
                    }
                    AddBrokenToken(lineAdmission, ref line, run, paintToken, logicalPaintToken, width, visibleTokenStart);
                    continue;
                }

                if (!preventTokenWrapping && canBreakBefore && line.HasFlowContent
                    && previewAdvance > line.ResolveAvailableWidth(width) + 0.0001D) {
                    TrimTrailingWhitespace(line);
                    lineAdmission.Add(line);
                    line = new InlineLine();
                    if (whitespace && !preserveWhitespace) continue;
                }

                if (hasTabs) {
                    AddTabExpandedSegments(line, normalizedToken, normalizedLogicalToken, run, visibleTokenStart);
                    continue;
                }
                line.Add(new InlineSegment(paintToken, measured, run, logicalPaintToken, logicalEndProgress: tokenEnd));
            }
        }

        if (noWrapRangeStart >= 0) {
            FinalizeNoWrapRange(
                lineAdmission,
                ref line,
                noWrapRangeStart,
                noWrapStrutStart,
                noWrapRangeStartedAfterContent,
                width);
        }

        TrimTrailingWhitespace(line);
        if (line.Segments.Count > 0 || lines.Count == 0) lineAdmission.Add(line);
        } catch (InlineLineLimitException) {
            onLineLimitExceeded?.Invoke();
        }
        ExpandLeaderSegments(lines, width);
        int completeLogicalProgress = lines
            .SelectMany(candidate => candidate.Segments)
            .Select(segment => segment.LogicalEndProgress)
            .DefaultIfEmpty(canonicalProgress)
            .Max();
        if (paragraphStyle.LineClamp.HasValue && lines.Count > paragraphStyle.LineClamp.Value) {
            lines.RemoveRange(paragraphStyle.LineClamp.Value, lines.Count - paragraphStyle.LineClamp.Value);
            ApplyEndEllipsis(lines[lines.Count - 1], width, completeLogicalProgress);
        } else if (paragraphStyle.TextOverflow == "ellipsis"
            && paragraphStyle.OverflowX != "visible") {
            foreach (InlineLine overflowingLine in lines.Where(candidate =>
                         candidate.Width > candidate.ResolveAvailableWidth(width) + 0.0001D)) {
                int lineLogicalProgress = overflowingLine.Segments
                    .Select(segment => segment.LogicalEndProgress)
                    .DefaultIfEmpty(completeLogicalProgress)
                    .Max();
                ApplyEndEllipsis(overflowingLine, width, lineLogicalProgress);
            }
        }
        return RenderInlineLines(
            lines,
            width,
            paragraphStyle,
            formattingContainer,
            supportsContinuationReflow: supportsContinuationReflow,
            isInlineContinuation: skipLogicalCharacters > 0,
            paintCapture: paintCapture);
    }

    private void ExpandLeaderSegments(IEnumerable<InlineLine> lines, double width) {
        foreach (InlineLine line in lines) {
            List<int> leaders = line.Segments
                .Select((segment, index) => new { segment, index })
                .Where(item => item.segment.Run.LeaderPattern != null)
                .Select(item => item.index)
                .ToList();
            if (leaders.Count == 0) continue;
            double availableWidth = line.ResolveAvailableWidth(width);
            double share = Math.Max(0D, availableWidth - line.Width) / leaders.Count;
            foreach (int index in leaders) line.SetSegmentWidth(index, share);
        }
    }

}
