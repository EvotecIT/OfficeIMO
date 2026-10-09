using AngleSharp.Dom;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private void AddFloatingRun(
        IElement element,
        double containingWidth,
        HtmlRenderBoxStyle parentStyle,
        int depth,
        HtmlRenderBoxStyle style,
        string? link,
        ICollection<HtmlInlineRun> runs) {
        double measuredOuterWidth = IsReplacedImageElement(element)
            ? ResolveFloatingImageOuterWidth(element, style)
            : ResolvePositionedOuterWidth(element, style, containingWidth, null, null, depth);
        double outerWidth = Math.Min(Math.Max(1D, containingWidth), measuredOuterWidth);
        HtmlRenderBoxStyle floatStyle = style.Clone();
        if (!floatStyle.ExplicitWidth.HasValue) SetPositionedExplicitWidth(floatStyle, outerWidth);
        floatStyle.FloatSide = "none";
        floatStyle.ClearSide = "none";
        floatStyle.UnsupportedFloat = string.Empty;
        floatStyle.UnsupportedClear = string.Empty;
        HtmlRenderFlowBlock block = LayoutElementWithoutEditableRegionMarker(
            element, outerWidth, floatStyle, parentStyle, depth + 1);
        block = WrapEditableLayoutRegion(block, element, style, HtmlRenderLayoutRegionKind.Floating);
        runs.Add(new HtmlInlineRun(
            block,
            style,
            link,
            HtmlRenderStyleResolver.DescribeSource(element),
            style.FloatSide,
            style.ClearSide,
            element));
    }

    private bool ContainsFloatingDescendant(IElement element, double containingWidth, HtmlRenderBoxStyle parentStyle, int depth) {
        EnsureDepth(depth, element);
        if (_containsInFlowFloatCache.TryGetValue(element, out bool cached)) return cached;
        foreach (IElement child in element.Children) {
            if (ShouldSkipElement(child)) continue;
            HtmlRenderBoxStyle style = _styleResolver.Resolve(child, containingWidth, parentStyle);
            if (style.Display == "none" || ShouldExtractOutOfFlow(style)) continue;
            if (style.FloatSide != "none" || ContainsFloatingDescendant(child, containingWidth, style, depth + 1)) {
                _containsInFlowFloatCache[element] = true;
                return true;
            }
        }
        _containsInFlowFloatCache[element] = false;
        return false;
    }

    private void ReportUnsupportedFloatValues(IElement element, HtmlRenderBoxStyle style) {
        if (style.UnsupportedFloat.Length == 0 && style.UnsupportedClear.Length == 0) return;
        if (!_reportedFloatValueFallbacks.Add(element)) return;
        var details = new List<string>(2);
        if (style.UnsupportedFloat.Length > 0) details.Add("float=" + style.UnsupportedFloat);
        if (style.UnsupportedClear.Length > 0) details.Add("clear=" + style.UnsupportedClear);
        _diagnostics.Add(
            ComponentName,
            HtmlRenderDiagnosticCodes.FloatValueUnsupported,
            "A float or clear property value used the normal-flow fallback.",
            HtmlDiagnosticSeverity.Warning,
            HtmlRenderStyleResolver.DescribeSource(element),
            string.Join(";", details));
    }

    private HtmlInlineLayout LayoutInlineRunsWithFloats(
        IReadOnlyList<HtmlInlineRun> runs,
        double width,
        HtmlRenderBoxStyle paragraphStyle,
        IElement? formattingContainer,
        PagedFloatBoundary? pageBoundary = null,
        IReadOnlyList<HtmlFloatExclusion>? inheritedFloats = null,
        InlinePaintCapture? paintCapture = null,
        InlineFloatContext? sharedContext = null) {
        var context = sharedContext ?? new InlineFloatContext(width, inheritedFloats);
        context.FirstLineIndent = paragraphStyle.TextIndent?.Resolve(width) ?? 0D;
        context.IndentAtRight = paragraphStyle.Direction == "rtl";
        var placements = new List<InlineFloatPlacement>();
        var lines = new List<InlineLine>();
        double y = 0D;
        InlineLine line = CreateFloatLine(context, ref y, paragraphStyle.LineHeight);
        bool previousWasCollapsibleSpace = false;
        int noWrapRangeStart = -1;
        bool noWrapRangeStartedAfterContent = false;

        for (int runIndex = 0; runIndex < runs.Count; runIndex++) {
            HtmlInlineRun run = runs[runIndex];
            if (ProcessInlineEdgeBoundary(run, line)) continue;
            if (run.FloatingBlock != null) {
                if (noWrapRangeStart >= 0) {
                    previousWasCollapsibleSpace = FinalizeFloatNoWrapRange(
                        lines,
                        ref line,
                        ref y,
                        context,
                        paragraphStyle.LineHeight,
                        noWrapRangeStart,
                        noWrapRangeStartedAfterContent);
                    noWrapRangeStart = -1;
                    noWrapRangeStartedAfterContent = false;
                }
                HtmlRenderFlowBlock floatingBlock = run.FloatingBlock;
                if (run.ClearSide == "none"
                    && pageBoundary.HasValue
                    && pageBoundary.Value.ShouldDefer(y, floatingBlock.Height)) {
                    InlineFloatPlacement deferred = context.Place(run, pageBoundary.Value.RemainingHeight);
                    placements.Add(deferred);
                    double available = pageBoundary.Value.RemainingHeight - y;
                    if (!floatingBlock.AvoidBreakInside
                        && floatingBlock.AvoidBreakRanges.Any(range =>
                            range.Start < available - 0.0001D && range.End > available + 0.0001D)
                        && FindFragmentEnd(floatingBlock, 0D, available,
                            fullPageHeight: pageBoundary.Value.PageHeight) <= 0.01D + 0.0001D) {
                        context.ReserveOriginatingPageExclusion(deferred, y, pageBoundary.Value.RemainingHeight);
                    }
                    _pagedFloatDeferredInRelayout = true;
                    continue;
                }
                InlineFloatBand floatBand = context.ResolveBand(y, floatingBlock.Height);
                bool sharesCurrentLine = line.HasFlowContent
                    && run.ClearSide == "none"
                    && Math.Abs(floatBand.Left - line.X) <= 0.0001D
                    && Math.Abs(floatBand.Width - line.AvailableWidth) <= 0.0001D
                    && line.Width + Math.Max(0.01D, floatingBlock.Width) <= floatBand.Width + 0.0001D;
                if (!sharesCurrentLine && line.Segments.Count > 0) {
                    CommitFloatLine(lines, ref line, ref y, context, paragraphStyle.LineHeight);
                }
                InlineFloatPlacement placement = context.Place(run, y);
                placements.Add(placement);
                if (sharesCurrentLine) {
                    InlineFloatBand remainingBand = context.ResolveBand(y, paragraphStyle.LineHeight);
                    line.Place(remainingBand.Left, y, remainingBand.Width);
                } else {
                    line = CreateFloatLine(context, ref y, paragraphStyle.LineHeight);
                }
                if (!sharesCurrentLine) previousWasCollapsibleSpace = false;
                continue;
            }
            bool runPreventsWrapping = !paragraphStyle.PreventTextWrapping && run.Style.PreventTextWrapping;
            if (!runPreventsWrapping && noWrapRangeStart >= 0) {
                previousWasCollapsibleSpace = FinalizeFloatNoWrapRange(
                    lines,
                    ref line,
                    ref y,
                    context,
                    paragraphStyle.LineHeight,
                    noWrapRangeStart,
                    noWrapRangeStartedAfterContent);
                noWrapRangeStart = -1;
                noWrapRangeStartedAfterContent = false;
            } else if (runPreventsWrapping && noWrapRangeStart < 0) {
                noWrapRangeStart = line.Segments.Count;
                noWrapRangeStartedAfterContent = line.HasFlowContent;
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
                previousWasCollapsibleSpace = false;
                continue;
            }
            if (run.AtomicBlock != null) {
                previousWasCollapsibleSpace = false;
                double atomicWidth = run.AtomicBlock.Width;
                if (!paragraphStyle.PreventTextWrapping
                    && !runPreventsWrapping
                    && line.HasFlowContent
                    && line.PreviewAdvance(run, atomicWidth) > line.AvailableWidth) {
                    CommitFloatLine(lines, ref line, ref y, context, paragraphStyle.LineHeight);
                }
                if (!line.HasFlowContent && line.PreviewAdvance(run, atomicWidth) > line.AvailableWidth + 0.0001D) {
                    MoveFloatLineBelowObstruction(ref line, ref y, context, paragraphStyle.LineHeight, line.PreviewAdvance(run, atomicWidth));
                }
                line.Add(new InlineSegment(string.Empty, atomicWidth, run));
                continue;
            }

            int logicalOffset = 0;
            bool preserveWhitespace = run.Style.PreserveWhitespace;
            IReadOnlyList<string> tokens = Tokenize(run.Text, preserveWhitespace, run.Style.BreakSpaces).ToList();
            for (int tokenIndex = 0; tokenIndex < tokens.Count; tokenIndex++) {
                string token = tokens[tokenIndex];
                run.InlineTokenEndsRun = tokenIndex == tokens.Count - 1;
                string logicalToken = SliceLogicalToken(run, token, ref logicalOffset);
                if (token == "\u2028" || preserveWhitespace && (token == "\n" || token == "\r\n")) {
                    if (noWrapRangeStart >= 0) {
                        FinalizeFloatNoWrapRange(
                            lines,
                            ref line,
                            ref y,
                            context,
                            paragraphStyle.LineHeight,
                            noWrapRangeStart,
                            noWrapRangeStartedAfterContent);
                    }
                    CommitFloatLine(lines, ref line, ref y, context, paragraphStyle.LineHeight, includeEmpty: true);
                    previousWasCollapsibleSpace = false;
                    noWrapRangeStart = runPreventsWrapping ? 0 : -1;
                    noWrapRangeStartedAfterContent = false;
                    continue;
                }

                bool whitespace = IsWhitespaceToken(token);
                string normalizedToken = !preserveWhitespace && whitespace ? " " : token;
                bool followsCollapsibleSpace = !whitespace && !preserveWhitespace && previousWasCollapsibleSpace;
                if (!preserveWhitespace && whitespace) {
                    if (!line.HasFlowContent || previousWasCollapsibleSpace) continue;
                    previousWasCollapsibleSpace = true;
                } else {
                    previousWasCollapsibleSpace = false;
                }

                bool hasTabs = preserveWhitespace && normalizedToken.IndexOf('\t') >= 0;
                double tabExpandedWidth = hasTabs ? MeasureTabExpandedText(normalizedToken, run.Style, line.PreviewContentStart(run)) : 0D;
                string expandedToken = hasTabs ? normalizedToken.Replace("\t", string.Empty) : normalizedToken;
                string normalizedLogicalToken = !preserveWhitespace && whitespace ? " " : logicalToken;
                HyphenationToken hyphenation = run.PreparedHyphenation.HasValue
                    ? run.PreparedHyphenation.Value.SliceSource(logicalOffset - token.Length, normalizedLogicalToken.Length)
                    : PrepareHyphenationToken(expandedToken, normalizedLogicalToken, run.Style);
                string paintToken = hyphenation.PaintText;
                double measured = hasTabs ? tabExpandedWidth : MeasureInlineText(paintToken, run.Style);
                if (run.EndsFirstLine && tokenIndex == tokens.Count - 1) {
                    paintToken += run.FirstLineHyphen;
                    double prefixWidth = MeasureInlineText(paintToken, run.Style);
                    if (line.HasFlowContent && line.PreviewAdvance(run, prefixWidth) > line.AvailableWidth + 0.0001D)
                        CommitFloatLine(lines, ref line, ref y, context, paragraphStyle.LineHeight);
                    if (!line.HasFlowContent && line.PreviewAdvance(run, prefixWidth) > line.AvailableWidth + 0.0001D)
                        MoveFloatLineBelowObstruction(ref line, ref y, context, paragraphStyle.LineHeight, line.PreviewAdvance(run, prefixWidth));
                    line.Add(new InlineSegment(paintToken, prefixWidth, run, hyphenation.LogicalText));
                    line.EndsWithHyphenation = run.FirstLineHyphen.Length > 0;
                    CommitFloatLine(lines, ref line, ref y, context, paragraphStyle.LineHeight);
                    continue;
                }
                bool preventTokenWrapping = paragraphStyle.PreventTextWrapping || runPreventsWrapping;
                if (!preventTokenWrapping
                    && !whitespace
                    && run.Style.WordBreak != "break-all"
                    && measured > Math.Max(0D, line.AvailableWidth - line.PreviewAdvance(run, 0D))
                    && TryAddHyphenatedFloatToken(
                        lines,
                        ref line,
                        ref y,
                        context,
                        paragraphStyle.LineHeight,
                        run,
                        hyphenation,
                        !HasRemainingInlineFlowContent(runs, runIndex, tokens, tokenIndex))) {
                    continue;
                }
                if (!preventTokenWrapping
                    && !whitespace
                    && measured > Math.Max(0D, line.AvailableWidth - line.PreviewAdvance(run, 0D))
                    && TryAddPreferredFloatBreakToken(
                        lines,
                        ref line,
                        ref y,
                        context,
                        paragraphStyle.LineHeight,
                        run,
                        hyphenation.PaintText,
                        hyphenation.LogicalText,
                        followsCollapsibleSpace)) {
                    continue;
                }
                bool breakAllIntoRemainingSpace = run.Style.WordBreak == "break-all"
                    && line.HasFlowContent
                    && measured > Math.Max(0D, line.AvailableWidth - line.PreviewAdvance(run, 0D));
                if (!preventTokenWrapping
                    && !whitespace
                    && AllowsEmergencyTokenBreak(run.Style)
                    && (InlineLine.PreviewEmptyLineAdvance(run, measured) > line.AvailableWidth || breakAllIntoRemainingSpace)) {
                    if (followsCollapsibleSpace && line.HasFlowContent && run.Style.WordBreak != "break-all") {
                        CommitFloatLine(lines, ref line, ref y, context, paragraphStyle.LineHeight);
                    }
                    AddBrokenFloatToken(lines, ref line, ref y, context, paragraphStyle.LineHeight, run, paintToken);
                    continue;
                }
                if (!preventTokenWrapping && line.HasFlowContent && line.PreviewAdvance(run, measured) > line.AvailableWidth) {
                    CommitFloatLine(lines, ref line, ref y, context, paragraphStyle.LineHeight);
                    if (whitespace && !preserveWhitespace) continue;
                }
                if (!preventTokenWrapping && !whitespace && line.PreviewAdvance(run, measured) > line.AvailableWidth && !line.HasFlowContent) {
                    MoveFloatLineBelowObstruction(ref line, ref y, context, paragraphStyle.LineHeight, line.PreviewAdvance(run, measured));
                }
                if (hasTabs) {
                    AddTabExpandedSegments(line, normalizedToken, normalizedToken, run);
                    continue;
                }
                line.Add(new InlineSegment(paintToken, measured, run, hyphenation.LogicalText));
            }
        }

        if (noWrapRangeStart >= 0) {
            FinalizeFloatNoWrapRange(
                lines,
                ref line,
                ref y,
                context,
                paragraphStyle.LineHeight,
                noWrapRangeStart,
                noWrapRangeStartedAfterContent);
        }

        TrimTrailingWhitespace(line);
        if (line.Segments.Count > 0) lines.Add(line);
        ExpandLeaderSegments(lines, width);
        if (paragraphStyle.LineClamp.HasValue && lines.Count > paragraphStyle.LineClamp.Value) {
            lines.RemoveRange(paragraphStyle.LineClamp.Value, lines.Count - paragraphStyle.LineClamp.Value);
            ApplyEndEllipsis(lines[lines.Count - 1], width, completeLogicalProgress: 0);
        } else if (paragraphStyle.TextOverflow == "ellipsis"
            && paragraphStyle.PreventTextWrapping
            && paragraphStyle.OverflowX != "visible"
            && lines.Count > 0
            && lines[0].Width > lines[0].AvailableWidth + 0.0001D) {
            ApplyEndEllipsis(lines[0], width, completeLogicalProgress: 0);
        }
        return RenderInlineLines(lines, width, paragraphStyle, formattingContainer, placements, sharedContext == null ? context.Bottom : 0D,
            paintCapture: paintCapture, flowExclusions: context.OriginatingPageExclusions,
            localFloatPageHeight: sharedContext == null && _options.Mode == HtmlRenderMode.Paged
                ? pageBoundary?.PageHeight ?? _activePageGeometry.ContentHeight : 0D);
    }

    private static bool FinalizeFloatNoWrapRange(
        ICollection<InlineLine> lines,
        ref InlineLine line,
        ref double y,
        InlineFloatContext context,
        double lineHeight,
        int rangeStart,
        bool startedAfterContent) {
        if (rangeStart < line.Segments.Count && line.Width > line.AvailableWidth + 0.0001D) {
            var range = line.Segments.Skip(rangeStart).ToArray();
            double rangeWidth = range.Sum(static segment => segment.Advance);
            bool canClearObstruction = context.NextBottomAfter(y) > y + 0.0001D;
            if (startedAfterContent || canClearObstruction) {
                while (line.Segments.Count > rangeStart) line.RemoveAt(rangeStart, preserveScopeContent: true);
                CommitFloatLine(lines, ref line, ref y, context, lineHeight);
                if (rangeWidth > line.AvailableWidth + 0.0001D) {
                    MoveFloatLineBelowObstruction(ref line, ref y, context, lineHeight, rangeWidth);
                }
                foreach (InlineSegment segment in range) {
                    if (!line.HasFlowContent
                        && segment.Run.RunningStringElement == null
                        && segment.Run.RunningElementAssignment == null
                        && IsWhitespaceToken(segment.Text)
                        && !segment.Run.Style.PreserveWhitespace) {
                        continue;
                    }
                    line.Add(segment);
                }
            }
        }

        for (int index = line.Segments.Count - 1; index >= 0; index--) {
            InlineSegment segment = line.Segments[index];
            if (segment.Run.RunningStringElement != null || segment.Run.RunningElementAssignment != null) continue;
            return IsWhitespaceToken(segment.Text) && !segment.Run.Style.PreserveWhitespace;
        }
        return false;
    }

    private void AddBrokenFloatToken(
        ICollection<InlineLine> lines,
        ref InlineLine line,
        ref double y,
        InlineFloatContext context,
        double lineHeight,
        HtmlInlineRun run,
        string token,
        bool completesToken = true) {
        int consumed = 0;
        foreach (string value in OfficeTextElements.Enumerate(token)) {
            consumed += value.Length;
            double elementWidth = MeasureInlineText(value, run.Style);
            if (line.HasFlowContent && line.PreviewAdvance(run, elementWidth, completesToken && consumed == token.Length) > line.AvailableWidth) {
                CommitFloatLine(lines, ref line, ref y, context, lineHeight);
            }
            line.Add(new InlineSegment(value, elementWidth, run));
        }
    }

    private static void CommitFloatLine(
        ICollection<InlineLine> lines,
        ref InlineLine line,
        ref double y,
        InlineFloatContext context,
        double defaultLineHeight,
        bool includeEmpty = false) {
        TrimTrailingWhitespace(line);
        if (line.Segments.Count > 0 || includeEmpty) {
            lines.Add(line);
            y = line.Y + line.ResolveLineHeight(defaultLineHeight);
            if (line.HasFlowContent || includeEmpty) context.FirstLineIndent = 0D;
        }
        line = CreateFloatLine(context, ref y, defaultLineHeight);
    }

    private static InlineLine CreateFloatLine(InlineFloatContext context, ref double y, double lineHeight) {
        InlineFloatBand band = context.ResolveUsableBand(ref y, lineHeight);
        var line = new InlineLine();
        line.Place(band.Left, y, band.Width);
        line.Indent(context.FirstLineIndent, context.IndentAtRight);
        return line;
    }

    private static void MoveFloatLineBelowObstruction(
        ref InlineLine line,
        ref double y,
        InlineFloatContext context,
        double lineHeight,
        double requiredWidth) {
        InlineSegment[] runningStringMarkers = line.Segments
            .Where(segment => segment.Run.RunningStringElement != null || segment.Run.RunningElementAssignment != null)
            .ToArray();
        double next = context.NextBottomAfter(y);
        while (next > y + 0.0001D) {
            y = next;
            InlineFloatBand band = context.ResolveUsableBand(ref y, lineHeight);
            line = new InlineLine();
            line.Place(band.Left, y, band.Width);
            line.Indent(context.FirstLineIndent, context.IndentAtRight);
            foreach (InlineSegment marker in runningStringMarkers) line.Add(marker);
            if (requiredWidth <= line.AvailableWidth + 0.0001D) return;
            next = context.NextBottomAfter(y);
        }
    }

    private HtmlRenderVisual CreateLeaderVisual(
        HtmlInlineRun run,
        double x,
        double y,
        double width,
        double lineHeight,
        int paintOrder) {
        HtmlRenderVisual paint;
        if (run.LeaderPattern == "." || run.LeaderPattern == "_") {
            double lineY = Math.Min(Math.Max(0.5D, lineHeight * 0.72D), Math.Max(0.5D, lineHeight - 0.5D));
            OfficeShape shape = OfficeShape.Line(0D, 0D, Math.Max(0.01D, width), 0D);
            shape.FillColor = null;
            shape.StrokeColor = run.Style.Color;
            shape.StrokeWidth = Math.Max(0.75D, run.Style.Font.Size / 16D);
            shape.Height = shape.StrokeWidth;
            shape.StrokeDashStyle = run.LeaderPattern == "." ? OfficeStrokeDashStyle.Dot : OfficeStrokeDashStyle.Solid;
            paint = new HtmlRenderShape(shape, x, y + lineY, paintOrder, run.LinkUri, run.Source);
        } else {
            string leaderPattern = run.LeaderPattern!;
            double unitWidth = Math.Max(0.01D, MeasureInlineText(leaderPattern, run.Style));
            double requiredCount = Math.Max(1D, Math.Ceiling(width / unitWidth));
            if (double.IsNaN(requiredCount) || double.IsInfinity(requiredCount)
                || requiredCount * leaderPattern.Length > _options.MaxLeaderCharacters) {
                throw new HtmlDomLimitException(
                    HtmlRenderDiagnosticCodes.LayoutOperationLimitExceeded,
                    "Generated leader text exceeded the configured character limit.",
                    nameof(HtmlRenderOptions.MaxLeaderCharacters),
                    (long)Math.Min(long.MaxValue, requiredCount * leaderPattern.Length),
                    _options.MaxLeaderCharacters);
            }
            int count = (int)requiredCount;
            ChargeLayoutOperations((long)count * leaderPattern.Length, "leader text expansion");
            string repeated = string.Concat(Enumerable.Repeat(leaderPattern, count));
            paint = new HtmlRenderText(
                repeated,
                x,
                y,
                Math.Max(0.01D, width),
                Math.Max(0.01D, lineHeight),
                run.Style.Font,
                run.Style.Color,
                OfficeTextAlignment.Left,
                lineHeight,
                paintOrder,
                run.LinkUri,
                run.Source,
                semanticRole: null,
                layoutY: null,
                semanticNodeId: null,
                textAdvanceWidth: Math.Max(0.01D, count * unitWidth),
                textPaintWidth: Math.Max(0.01D, count * unitWidth),
                featureSettings: run.Style.TextFeatureSettings,
                fontPalette: run.Style.FontPalette,
                fontDescriptor: run.Style.FontDescriptor);
        }

        return new HtmlRenderSemanticGroup(
            HtmlRenderSemanticGroupRole.Artifact,
            x,
            y,
            Math.Max(0.01D, width),
            Math.Max(0.01D, lineHeight),
            new[] { paint },
            paintOrder,
            run.Source);
    }

    private readonly struct PagedFloatBoundary {
        internal PagedFloatBoundary(double remainingHeight, double pageHeight, double fragmentStart = 0D) {
            RemainingHeight = remainingHeight;
            PageHeight = pageHeight;
            FragmentStart = fragmentStart;
        }

        internal double RemainingHeight { get; }
        internal double PageHeight { get; }
        internal double FragmentStart { get; }

        internal PagedFloatBoundary Shift(double offset) => new PagedFloatBoundary(RemainingHeight - offset, PageHeight, FragmentStart - offset);

        internal bool ShouldDefer(double y, double floatHeight) =>
            RemainingHeight > 0.0001D
            && y < RemainingHeight - 0.0001D
            && y + floatHeight > RemainingHeight + 0.0001D
            && floatHeight <= PageHeight + 0.0001D;
    }

    private sealed class InlineFloatPlacement {
        internal InlineFloatPlacement(HtmlInlineRun run, double x, double y, double width, double height) {
            Run = run;
            X = x;
            Y = y;
            Width = width;
            Height = height;
        }

        internal HtmlInlineRun Run { get; }
        internal double X { get; }
        internal double Y { get; }
        internal double Width { get; }
        internal double Height { get; }
        internal double Right => X + Width;
        internal double Bottom => Y + Height;
    }

    private readonly struct InlineFloatBand {
        internal InlineFloatBand(double left, double right) {
            Left = left;
            Right = Math.Max(left, right);
        }

        internal double Left { get; }
        internal double Right { get; }
        internal double Width => Math.Max(0.01D, Right - Left);
    }
}
