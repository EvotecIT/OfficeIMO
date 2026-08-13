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
        double measuredOuterWidth = string.Equals(element.TagName, "img", StringComparison.OrdinalIgnoreCase)
            ? ResolveFloatingImageOuterWidth(element, style)
            : ResolvePositionedOuterWidth(element, style, containingWidth, null, null);
        double outerWidth = Math.Min(Math.Max(1D, containingWidth), measuredOuterWidth);
        HtmlRenderBoxStyle floatStyle = style.Clone();
        if (!floatStyle.ExplicitWidth.HasValue) SetPositionedExplicitWidth(floatStyle, outerWidth);
        floatStyle.FloatSide = "none";
        floatStyle.ClearSide = "none";
        floatStyle.UnsupportedFloat = string.Empty;
        floatStyle.UnsupportedClear = string.Empty;
        HtmlRenderFlowBlock block = LayoutElement(element, outerWidth, floatStyle, parentStyle, depth + 1);
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
        IElement? formattingContainer) {
        var context = new InlineFloatContext(width);
        var placements = new List<InlineFloatPlacement>();
        var lines = new List<InlineLine>();
        double y = 0D;
        InlineLine line = CreateFloatLine(context, ref y, paragraphStyle.LineHeight);
        bool previousWasCollapsibleSpace = false;
        int noWrapRangeStart = -1;
        bool noWrapRangeStartedAfterContent = false;

        foreach (HtmlInlineRun run in runs) {
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
                if (line.Segments.Count > 0) CommitFloatLine(lines, ref line, ref y, context, paragraphStyle.LineHeight);
                InlineFloatPlacement placement = context.Place(run, y);
                placements.Add(placement);
                line = CreateFloatLine(context, ref y, paragraphStyle.LineHeight);
                previousWasCollapsibleSpace = false;
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
            if (run.PositionedMarkerElement != null) {
                line.Add(new InlineSegment(string.Empty, 0D, run));
                previousWasCollapsibleSpace = false;
                continue;
            }
            if (run.AtomicBlock != null) {
                previousWasCollapsibleSpace = false;
                double atomicWidth = run.AtomicBlock.Width;
                if (!line.HasFlowContent && atomicWidth > line.AvailableWidth + 0.0001D) {
                    MoveFloatLineBelowObstruction(ref line, ref y, context, paragraphStyle.LineHeight, atomicWidth);
                }
                if (!paragraphStyle.PreventTextWrapping
                    && !runPreventsWrapping
                    && line.HasFlowContent
                    && line.Width + atomicWidth > line.AvailableWidth) {
                    CommitFloatLine(lines, ref line, ref y, context, paragraphStyle.LineHeight);
                }
                line.Add(new InlineSegment(string.Empty, atomicWidth, run));
                continue;
            }

            bool preserveWhitespace = run.Style.PreserveWhitespace;
            foreach (string token in Tokenize(run.Text, preserveWhitespace, run.Style.BreakSpaces)) {
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
                if (!preserveWhitespace && whitespace) {
                    if (!line.HasFlowContent || previousWasCollapsibleSpace) continue;
                    previousWasCollapsibleSpace = true;
                } else {
                    previousWasCollapsibleSpace = false;
                }

                bool hasTabs = preserveWhitespace && normalizedToken.IndexOf('\t') >= 0;
                double tabExpandedWidth = hasTabs ? MeasureTabExpandedText(normalizedToken, run.Style, line.Width) : 0D;
                string expandedToken = hasTabs ? normalizedToken.Replace("\t", string.Empty) : normalizedToken;
                HyphenationToken hyphenation = PrepareHyphenationToken(expandedToken, normalizedToken, run.Style);
                string paintToken = hyphenation.PaintText;
                double measured = hasTabs ? tabExpandedWidth : MeasureInlineText(paintToken, run.Style);
                bool preventTokenWrapping = paragraphStyle.PreventTextWrapping || runPreventsWrapping;
                if (!preventTokenWrapping
                    && !whitespace
                    && run.Style.WordBreak != "break-all"
                    && measured > Math.Max(0D, line.AvailableWidth - line.Width)
                    && TryAddHyphenatedFloatToken(
                        lines,
                        ref line,
                        ref y,
                        context,
                        paragraphStyle.LineHeight,
                        run,
                        hyphenation)) {
                    continue;
                }
                if (!preventTokenWrapping
                    && !whitespace
                    && measured > Math.Max(0D, line.AvailableWidth - line.Width)
                    && TryAddPreferredFloatBreakToken(
                        lines,
                        ref line,
                        ref y,
                        context,
                        paragraphStyle.LineHeight,
                        run,
                        hyphenation.PaintText,
                        hyphenation.LogicalText)) {
                    continue;
                }
                bool breakAllIntoRemainingSpace = run.Style.WordBreak == "break-all"
                    && line.HasFlowContent
                    && measured > Math.Max(0D, line.AvailableWidth - line.Width);
                if (!preventTokenWrapping
                    && !whitespace
                    && AllowsEmergencyTokenBreak(run.Style)
                    && (measured > line.AvailableWidth || breakAllIntoRemainingSpace)) {
                    AddBrokenFloatToken(lines, ref line, ref y, context, paragraphStyle.LineHeight, run, paintToken);
                    continue;
                }
                if (!preventTokenWrapping && !whitespace && measured > line.AvailableWidth && !line.HasFlowContent) {
                    MoveFloatLineBelowObstruction(ref line, ref y, context, paragraphStyle.LineHeight, measured);
                }
                if (!preventTokenWrapping && line.HasFlowContent && line.Width + measured > line.AvailableWidth) {
                    CommitFloatLine(lines, ref line, ref y, context, paragraphStyle.LineHeight);
                    if (whitespace && !preserveWhitespace) continue;
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
        return RenderInlineLines(lines, width, paragraphStyle, formattingContainer, placements, context.Bottom);
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
            double rangeWidth = range.Sum(static segment => segment.Width);
            bool canClearObstruction = context.NextBottomAfter(y) > y + 0.0001D;
            if (startedAfterContent || canClearObstruction) {
                while (line.Segments.Count > rangeStart) line.RemoveAt(rangeStart);
                CommitFloatLine(lines, ref line, ref y, context, lineHeight);
                if (rangeWidth > line.AvailableWidth + 0.0001D) {
                    MoveFloatLineBelowObstruction(ref line, ref y, context, lineHeight, rangeWidth);
                }
                foreach (InlineSegment segment in range) {
                    if (!line.HasFlowContent
                        && segment.Run.RunningStringElement == null
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
            if (segment.Run.RunningStringElement != null) continue;
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
        string token) {
        foreach (string value in OfficeTextElements.Enumerate(token)) {
            double elementWidth = MeasureInlineText(value, run.Style);
            if (line.HasFlowContent && line.Width + elementWidth > line.AvailableWidth) {
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
        }
        line = CreateFloatLine(context, ref y, defaultLineHeight);
    }

    private static InlineLine CreateFloatLine(InlineFloatContext context, ref double y, double lineHeight) {
        InlineFloatBand band = context.ResolveUsableBand(ref y, lineHeight);
        var line = new InlineLine();
        line.Place(band.Left, y, band.Width);
        return line;
    }

    private static void MoveFloatLineBelowObstruction(
        ref InlineLine line,
        ref double y,
        InlineFloatContext context,
        double lineHeight,
        double requiredWidth) {
        InlineSegment[] runningStringMarkers = line.Segments
            .Where(segment => segment.Run.RunningStringElement != null)
            .ToArray();
        double next = context.NextBottomAfter(y);
        while (next > y + 0.0001D) {
            y = next;
            InlineFloatBand band = context.ResolveUsableBand(ref y, lineHeight);
            line = new InlineLine();
            line.Place(band.Left, y, band.Width);
            foreach (InlineSegment marker in runningStringMarkers) line.Add(marker);
            if (requiredWidth <= band.Width + 0.0001D) return;
            next = context.NextBottomAfter(y);
        }
    }

    private HtmlInlineLayout RenderInlineLines(
        IReadOnlyList<InlineLine> lines,
        double width,
        HtmlRenderBoxStyle paragraphStyle,
        IElement? formattingContainer,
        IReadOnlyList<InlineFloatPlacement>? floatPlacements = null,
        double minimumHeight = 0D,
        bool supportsContinuationReflow = false) {
        var visuals = new List<HtmlRenderVisual>();
        var ownedVisuals = new Dictionary<IElement, List<HtmlRenderVisual>>();
        var inlineBounds = new Dictionary<IElement, InlineContainingBounds>();
        var breakOffsets = new SortedSet<double>();
        var breakProgress = new List<HtmlInlineBreakProgress>();
        var runningStringAssignments = new List<HtmlCssRunningStringAssignment>();
        IReadOnlyList<IReadOnlyList<InlineSegment>> mergedLines = lines
            .Select(static line => (IReadOnlyList<InlineSegment>)MergeAdjacentInlineSegments(line.Segments))
            .ToArray();
        IReadOnlyList<IReadOnlyList<InlineSegment>>? resolvedBidiLines = ResolveInlineParagraphSegments(
            mergedLines,
            paragraphStyle.Direction);

        if (floatPlacements != null) {
            foreach (InlineFloatPlacement placement in floatPlacements) {
                HtmlInlineRun run = placement.Run;
                HtmlRenderFlowBlock block = run.FloatingBlock!;
                RecordInlineOwnerGeometry(run, formattingContainer, placement.X, placement.Y, placement.Width, placement.Height, inlineBounds);
                foreach (HtmlCssRunningStringAssignment assignment in block.RunningStringAssignments) {
                    runningStringAssignments.Add(assignment.Translate(placement.Y));
                }
                if (run.Style.PaintVisible) {
                    foreach (HtmlRenderVisual visual in block.Visuals) {
                        AddInlineOwnedVisual(
                            visuals,
                            ownedVisuals,
                            visual.Translate(placement.X, placement.Y, visuals.Count),
                            run.OwnerElement,
                            formattingContainer);
                    }
                    if (run.LinkUri != null) {
                        OfficeShape linkArea = OfficeShape.Rectangle(placement.Width, placement.Height);
                        linkArea.FillColor = null;
                        linkArea.StrokeWidth = 0D;
                        AddInlineOwnedVisual(
                            visuals,
                            ownedVisuals,
                            new HtmlRenderShape(linkArea, placement.X, placement.Y, visuals.Count, run.LinkUri, run.Source),
                            run.OwnerElement,
                            formattingContainer);
                    }
                }
            }
        }

        double flowY = 0D;
        int logicalCharacters = 0;
        for (int lineIndex = 0; lineIndex < lines.Count; lineIndex++) {
            InlineLine current = lines[lineIndex];
            double lineHeight = current.ResolveLineHeight(paragraphStyle.LineHeight);
            double baseline = current.ResolveBaseline(paragraphStyle.LineHeight);
            double lineY = current.HasExplicitPlacement ? current.Y : flowY;
            double availableWidth = current.HasExplicitPlacement ? current.AvailableWidth : width;
            double lineX = current.HasExplicitPlacement ? current.X : 0D;
            double offsetX = ResolveLineOffset(paragraphStyle.Alignment, availableWidth, current.Width);
            double lineStart = lineX + offsetX;
            double lineRight = lineX + availableWidth;
            IReadOnlyList<InlineSegment> paintLineSegments = resolvedBidiLines?[lineIndex] ?? mergedLines[lineIndex];
            bool lineBidiResolved = resolvedBidiLines != null;
            bool rightToLeftLine = !lineBidiResolved
                && string.Equals(paragraphStyle.Direction, "rtl", StringComparison.Ordinal)
                && current.Segments.Any(segment => OfficeTextElements.ContainsRightToLeft(segment.Text));
            double cursor = rightToLeftLine ? lineStart + current.Width : lineStart;
            int lineVisualStart = visuals.Count;
            foreach (InlineSegment segment in paintLineSegments) {
                double x = rightToLeftLine ? cursor - segment.Width : cursor;
                if (segment.Run.RunningStringElement != null) {
                    runningStringAssignments.AddRange(ResolveRunningStringAssignments(
                        segment.Run.RunningStringElement,
                        segment.Run.Style,
                        lineY));
                } else if (segment.Run.PositionedMarkerElement != null) {
                    RecordInlineStaticMarker(segment.Run, formattingContainer, x, lineY, lineHeight, inlineBounds);
                    EnsureInlineStackingOwner(segment.Run.OwnerElement, formattingContainer, ownedVisuals);
                } else if (segment.Run.AtomicBlock != null) {
                    HtmlRenderFlowBlock atomic = segment.Run.AtomicBlock;
                    double atomicBaseline = segment.Run.AtomicBaseline ?? atomic.Height;
                    double atomicY = lineY + Math.Max(0D, (current.HasReplacedImage ? baseline : lineHeight) - atomicBaseline);
                    RecordInlineOwnerGeometry(segment.Run, formattingContainer, x, atomicY, segment.Width, atomic.Height, inlineBounds);
                    foreach (HtmlCssRunningStringAssignment assignment in atomic.RunningStringAssignments) {
                        runningStringAssignments.Add(assignment.Translate(atomicY));
                    }
                    if (segment.Run.Style.PaintVisible) {
                        foreach (HtmlRenderVisual visual in atomic.Visuals) {
                            HtmlRenderVisual translated = visual.Translate(x, atomicY, visuals.Count);
                            if (Math.Abs(segment.Run.PaintOffsetX) > 0.0001D || Math.Abs(segment.Run.PaintOffsetY) > 0.0001D) {
                                translated = translated.TranslatePaint(segment.Run.PaintOffsetX, segment.Run.PaintOffsetY, visuals.Count);
                            }
                            AddInlineOwnedVisual(visuals, ownedVisuals, translated, segment.Run.OwnerElement, formattingContainer);
                        }
                        if (segment.Run.LinkUri != null) {
                            OfficeShape linkArea = OfficeShape.Rectangle(Math.Max(0.01D, segment.Width), Math.Max(0.01D, atomic.Height));
                            linkArea.FillColor = null;
                            linkArea.StrokeWidth = 0D;
                            HtmlRenderVisual linkVisual = new HtmlRenderShape(linkArea, x, atomicY, visuals.Count, segment.Run.LinkUri, segment.Run.Source);
                            if (Math.Abs(segment.Run.PaintOffsetX) > 0.0001D || Math.Abs(segment.Run.PaintOffsetY) > 0.0001D) {
                                linkVisual = linkVisual.TranslatePaint(segment.Run.PaintOffsetX, segment.Run.PaintOffsetY, visuals.Count);
                            }
                            AddInlineOwnedVisual(visuals, ownedVisuals, linkVisual, segment.Run.OwnerElement, formattingContainer);
                        }
                    }
                } else if (segment.Text.Length > 0) {
                    double textLineHeight = current.HasReplacedImage ? segment.Run.Style.LineHeight : lineHeight;
                    double textY = current.HasReplacedImage
                        ? lineY + Math.Max(0D, baseline - ResolveTextAscent(segment.Run.Style))
                        : lineY;
                    RecordInlineOwnerGeometry(segment.Run, formattingContainer, x, textY, Math.Max(0.01D, segment.Width), textLineHeight, inlineBounds);
                    if (!segment.Run.Style.PaintVisible) {
                        cursor += rightToLeftLine ? -segment.Width : segment.Width;
                        continue;
                    }
                    double frameTolerance = Math.Max(1D, segment.Run.Style.Font.Size * 0.35D);
                    IReadOnlyList<InlinePaintSegment> paintSegments = ResolveInlinePaintSegments(segment, x);
                    var textVisuals = new List<HtmlRenderVisual>(paintSegments.Count);
                    foreach (InlinePaintSegment paintSegment in paintSegments) {
                        double frameWidth = Math.Min(Math.Max(0.01D, lineRight - paintSegment.X), paintSegment.Width + frameTolerance);
                        textVisuals.Add(new HtmlRenderText(
                            paintSegment.Text,
                            paintSegment.X,
                            textY,
                            Math.Max(0.01D, frameWidth),
                            Math.Max(0.01D, textLineHeight),
                            segment.Run.Style.Font,
                            segment.Run.Style.Color,
                            OfficeTextAlignment.Left,
                            textLineHeight,
                            textVisuals.Count,
                            segment.Run.LinkUri,
                            segment.Run.Source,
                            segment.Run.SemanticRole,
                            layoutY: null,
                            semanticNodeId: segment.Run.SemanticNodeId,
                            textAdvanceWidth: paintSegment.Advance,
                            bidiVisualOrderResolved: segment.BidiResolved));
                    }
                    HtmlRenderVisual textVisual = paintSegments.Count > 1 || segment.BidiResolved ||
                        !string.Equals(segment.Text, segment.LogicalText, StringComparison.Ordinal) ||
                        OfficeTextElements.ContainsRightToLeft(segment.Text) || OfficeTextElements.ContainsBidiControl(segment.Text)
                        ? new HtmlRenderLogicalTextGroup(
                            OfficeTextElements.ContainsBidiControl(segment.LogicalText)
                                ? string.Concat(OfficeBidiTextResolver.ResolveRuns(segment.LogicalText).Select(static run => run.Text))
                                : segment.LogicalText,
                            x,
                            textY,
                            Math.Max(0.01D, segment.Width),
                            Math.Max(0.01D, textLineHeight),
                            textVisuals,
                            visuals.Count,
                            segment.Run.Source)
                        : textVisuals[0];
                    AddInlineOwnedVisual(
                        visuals,
                        ownedVisuals,
                        textVisual.TranslatePaint(segment.Run.PaintOffsetX, segment.Run.PaintOffsetY, visuals.Count),
                        segment.Run.OwnerElement,
                        formattingContainer);
                }
                cursor += rightToLeftLine ? -segment.Width : segment.Width;
            }

            if (lineBidiResolved && visuals.Count > lineVisualStart) {
                List<HtmlRenderVisual> lineVisuals = visuals.GetRange(lineVisualStart, visuals.Count - lineVisualStart);
                visuals.RemoveRange(lineVisualStart, visuals.Count - lineVisualStart);
                visuals.Add(new HtmlRenderLogicalTextGroup(
                    ResolveRootInlineLineLogicalText(mergedLines[lineIndex], formattingContainer),
                    lineStart,
                    lineY,
                    Math.Max(0.01D, current.Width),
                    Math.Max(0.01D, lineHeight),
                    lineVisuals,
                    lineVisualStart,
                    paintLineSegments.FirstOrDefault()?.Run.Source));
            }

            flowY = Math.Max(flowY, lineY + lineHeight);
            logicalCharacters = Math.Max(
                logicalCharacters,
                mergedLines[lineIndex].Select(segment => segment.LogicalEndProgress).DefaultIfEmpty(logicalCharacters).Max());
            if (lineHeight > 0D) {
                double breakOffset = lineY + lineHeight;
                breakOffsets.Add(breakOffset);
                breakProgress.Add(new HtmlInlineBreakProgress(breakOffset, logicalCharacters, formattingContainer));
            }
        }

        double height = Math.Max(flowY, minimumHeight);
        if (height > 0D) breakOffsets.Add(height);
        return new HtmlInlineLayout(
            ComposeInlinePositionedVisuals(visuals, ownedVisuals, inlineBounds, formattingContainer),
            height,
            breakOffsets,
            runningStringAssignments.OrderBy(assignment => assignment.OrderOffset),
            breakProgress,
            supportsContinuationReflow);
    }

    private sealed class InlineFloatContext {
        private readonly double _width;
        private readonly List<InlineFloatPlacement> _placements = new List<InlineFloatPlacement>();

        internal InlineFloatContext(double width) {
            _width = Math.Max(1D, width);
        }

        internal double Bottom => _placements.Count == 0 ? 0D : _placements.Max(item => item.Bottom);

        internal InlineFloatPlacement Place(HtmlInlineRun run, double requestedY) {
            HtmlRenderFlowBlock block = run.FloatingBlock!;
            double boxWidth = Math.Min(_width, Math.Max(0.01D, block.Width));
            double boxHeight = Math.Max(0.01D, block.Height);
            double y = Math.Max(requestedY, Clearance(run.ClearSide));
            while (true) {
                InlineFloatBand band = ResolveBand(y, boxHeight);
                if (boxWidth <= band.Width + 0.0001D) {
                    double x = run.FloatSide == "right" ? band.Right - boxWidth : band.Left;
                    var placement = new InlineFloatPlacement(run, x, y, boxWidth, boxHeight);
                    _placements.Add(placement);
                    return placement;
                }
                double next = NextBottomAfter(y);
                if (next <= y + 0.0001D) {
                    double x = run.FloatSide == "right" ? Math.Max(0D, _width - boxWidth) : 0D;
                    var placement = new InlineFloatPlacement(run, x, y, boxWidth, boxHeight);
                    _placements.Add(placement);
                    return placement;
                }
                y = next;
            }
        }

        internal InlineFloatBand ResolveUsableBand(ref double y, double height) {
            InlineFloatBand band = ResolveBand(y, height);
            while (band.Width <= 0.01D) {
                double next = NextBottomAfter(y);
                if (next <= y + 0.0001D) break;
                y = next;
                band = ResolveBand(y, height);
            }
            return band;
        }

        internal InlineFloatBand ResolveBand(double y, double height) {
            double left = 0D;
            double right = _width;
            double bottom = y + Math.Max(0.01D, height);
            foreach (InlineFloatPlacement placement in _placements) {
                if (placement.Y >= bottom - 0.0001D || placement.Bottom <= y + 0.0001D) continue;
                if (placement.Run.FloatSide == "right") right = Math.Min(right, placement.X);
                else left = Math.Max(left, placement.Right);
            }
            return new InlineFloatBand(left, Math.Max(left, right));
        }

        internal double NextBottomAfter(double y) => _placements
            .Where(item => item.Bottom > y + 0.0001D)
            .Select(item => item.Bottom)
            .DefaultIfEmpty(y)
            .Min();

        private double Clearance(string clearSide) {
            if (clearSide == "none") return 0D;
            return _placements
                .Where(item => clearSide == "both" || item.Run.FloatSide == clearSide)
                .Select(item => item.Bottom)
                .DefaultIfEmpty(0D)
                .Max();
        }
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
