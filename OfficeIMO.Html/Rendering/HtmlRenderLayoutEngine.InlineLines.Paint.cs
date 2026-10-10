using AngleSharp.Dom;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private HtmlInlineLayout RenderInlineLines(
        IReadOnlyList<InlineLine> lines,
        double width,
        HtmlRenderBoxStyle paragraphStyle,
        IElement? formattingContainer,
        IReadOnlyList<InlineFloatPlacement>? floatPlacements = null,
        double minimumHeight = 0D,
        bool supportsContinuationReflow = false,
        bool isInlineContinuation = false,
        InlinePaintCapture? paintCapture = null, IReadOnlyList<HtmlFloatExclusion>? flowExclusions = null,
        double localFloatPageHeight = 0D) {
        var visuals = new List<HtmlRenderVisual>();
        var ownedVisuals = new Dictionary<IElement, List<HtmlRenderVisual>>();
        var inlineBounds = new Dictionary<IElement, InlineContainingBounds>();
        var anchorBounds = new Dictionary<IElement, InlineAnchorBounds>();
        var breakOffsets = new SortedSet<double>();
        var breakProgress = new List<HtmlInlineBreakProgress>();
        var runningStringAssignments = new List<HtmlCssRunningStringAssignment>();
        bool hasFloats = floatPlacements?.Count > 0;
        var forcedBreaks = hasFloats ? new List<HtmlRenderForcedBreak>() : null;
        var lineBreakGroups = hasFloats ? new List<HtmlRenderLineBreakGroup>() : null;
        var continuationGroups = hasFloats ? new List<HtmlRenderContinuationGroup>() : null;
        var trailingGroups = hasFloats ? new List<HtmlRenderTrailingGroup>() : null;
        _currentInlineEdgeGeometry.Clear();
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
                RecordInlineAnchorGeometry(run, formattingContainer, placement.X, placement.Y,
                    placement.Width, placement.Height, anchorBounds);
                foreach (HtmlCssRunningStringAssignment assignment in block.RunningStringAssignments) {
                    runningStringAssignments.Add(assignment.Translate(placement.Y));
                }
                AppendFloatFragmentationMetadata(block, placement, forcedBreaks!, lineBreakGroups!, continuationGroups!, trailingGroups!);
                if (run.Style.PaintVisible) {
                    foreach (HtmlRenderVisual visual in block.Visuals) {
                        HtmlRenderVisual floated = visual.Translate(placement.X, placement.Y, visuals.Count);
                        floated.PaintPhase = HtmlRenderPaintPhase.Float;
                        AddInlineOwnedVisual(
                            visuals,
                            ownedVisuals,
                            floated,
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
            bool alignTextBaseline = current.HasMixedTextSizes(paragraphStyle);
            bool emptyAtomicBaseline = TryResolveEmptyInlineBlockLineMetrics(current, paragraphStyle, ref lineHeight, out double resolvedEmptyBaseline);
            double baseline = emptyAtomicBaseline ? resolvedEmptyBaseline : alignTextBaseline || current.HasReplacedImage
                ? ResolveInlineSharedBaseline(current, paragraphStyle, ref lineHeight)
                : current.ResolveBaseline(paragraphStyle);
            double lineY = current.HasExplicitPlacement ? current.Y : flowY;
            double availableWidth = current.ResolveAvailableWidth(width);
            double lineX = (current.HasExplicitPlacement ? current.X : 0D) + current.IndentOffset;
            double offsetX = current.ResolveAlignmentOffset(paragraphStyle.Alignment, width);
            double lineStart = lineX + offsetX;
            double lineRight = lineX + availableWidth;
            IReadOnlyList<InlineSegment> paintLineSegments = resolvedBidiLines?[lineIndex] ?? mergedLines[lineIndex];
            bool lineBidiResolved = resolvedBidiLines != null;
            bool rightToLeftLine = !lineBidiResolved
                && string.Equals(paragraphStyle.Direction, "rtl", StringComparison.Ordinal)
                && current.Segments.Any(segment => OfficeTextElements.ContainsRightToLeft(segment.Text));
            double cursor = rightToLeftLine ? lineStart + current.Width : lineStart;
            if (lineBidiResolved) {
                _currentInlineEdgeGeometry.Clear();
                foreach (HtmlInlineEdgeScope scope in current.Segments.SelectMany(segment => segment.Run.InlineEdgeScopes).Distinct()) {
                    if (_reportedInlineEdgeFallbacks.Add(scope.Owner)) _diagnostics.Add(ComponentName,
                        HtmlRenderDiagnosticCodes.InlinePaintEffectUnsupported,
                        "Bidi-reordered inline edge fragments retain content-based decoration geometry; exact fragment edges are not qualified.",
                        HtmlDiagnosticSeverity.Warning, HtmlRenderStyleResolver.DescribeSource(scope.Owner),
                        "inline-edges;bidi-fragment-geometry", OfficeConversionLossKind.Approximation);
                }
            } else {
                current.RecordInlineScopePositions(lineStart, _currentInlineEdgeGeometry);
            }
            int lineVisualStart = visuals.Count;
            foreach (InlineSegment segment in paintLineSegments) {
                double x = rightToLeftLine ? cursor - segment.LeadingAdvance - segment.Width : cursor + segment.LeadingAdvance;
                if (segment.Run.RunningStringElement != null) {
                    runningStringAssignments.AddRange(ResolveRunningStringAssignments(
                        segment.Run.RunningStringElement,
                        segment.Run.Style,
                        lineY));
                } else if (segment.Run.RunningElementAssignment != null) {
                    runningStringAssignments.AddRange(segment.Run.RunningElementAssignments.Select(assignment => assignment.Translate(lineY)));
                } else if (segment.Run.PositionedMarkerElement != null) {
                    RecordInlineStaticMarker(segment.Run, formattingContainer, x, lineY, lineHeight, inlineBounds);
                    EnsureInlineStackingOwner(segment.Run.OwnerElement, formattingContainer, ownedVisuals);
                } else if (segment.Run.LeaderPattern != null && segment.Width > 0.0001D) {
                    RecordInlineOwnerGeometry(segment.Run, formattingContainer, x, lineY, segment.Width, lineHeight, inlineBounds);
                    RecordInlineAnchorGeometry(segment.Run, formattingContainer, x, lineY,
                        segment.Width, lineHeight, anchorBounds);
                    if (segment.Run.Style.PaintVisible && segment.Run.LeaderPattern != " ") {
                        HtmlRenderVisual leaderVisual = CreateLeaderVisual(segment.Run, x, lineY, segment.Width, lineHeight, visuals.Count);
                        AddInlineOwnedVisual(
                            visuals,
                            ownedVisuals,
                            leaderVisual.TranslateRelativePaint(segment.Run.PaintOffsetX, segment.Run.PaintOffsetY, leaderVisual.PaintOrder),
                            segment.Run.OwnerElement,
                            formattingContainer);
                    }
                } else if (segment.Run.IsEmptyInlineBox) {
                    RecordInlineOwnerGeometry(segment.Run, formattingContainer, x, lineY, 0D, lineHeight, inlineBounds);
                    RecordInlineAnchorGeometry(segment.Run, formattingContainer, x, lineY, 0D, lineHeight, anchorBounds);
                } else if (segment.Run.AtomicBlock != null) {
                    HtmlRenderFlowBlock atomic = segment.Run.AtomicBlock;
                    double atomicBaseline = segment.Run.AtomicBaseline ?? atomic.Height;
                    double atomicY = lineY + Math.Max(0D, (current.HasReplacedImage || emptyAtomicBaseline ? baseline : lineHeight) - atomicBaseline);
                    if (segment.Run.Style.TableVerticalAlignment == "top") atomicY = lineY;
                    RecordInlineOwnerGeometry(segment.Run, formattingContainer, x, atomicY, segment.Width, atomic.Height, inlineBounds);
                    double anchorTop = Math.Min(lineY, atomicY);
                    RecordInlineAnchorGeometry(segment.Run, formattingContainer, x, anchorTop,
                        segment.Width, Math.Max(lineY + lineHeight, atomicY + atomic.Height) - anchorTop,
                        anchorBounds);
                    foreach (HtmlCssRunningStringAssignment assignment in atomic.RunningStringAssignments) {
                        runningStringAssignments.Add(assignment.Translate(atomicY));
                    }
                    if (segment.Run.Style.PaintVisible) {
                        foreach (HtmlRenderVisual visual in atomic.Visuals) {
                            HtmlRenderVisual translated = visual.Translate(x, atomicY, visuals.Count);
                            if (Math.Abs(segment.Run.PaintOffsetX) > 0.0001D || Math.Abs(segment.Run.PaintOffsetY) > 0.0001D) {
                                translated = translated.TranslateRelativePaint(segment.Run.PaintOffsetX, segment.Run.PaintOffsetY, visuals.Count);
                            }
                            AddInlineOwnedVisual(visuals, ownedVisuals, ApplyInlineElementSemantics(translated, segment.Run), segment.Run.OwnerElement, formattingContainer);
                        }
                        if (segment.Run.SemanticNodeId.HasValue
                            && !string.IsNullOrWhiteSpace(segment.Run.BookmarkAnchorText)
                            && !ContainsBookmarkAnchor(atomic.Visuals, segment.Run.SemanticNodeId.Value)) {
                            AddInlineOwnedVisual(
                                visuals,
                                ownedVisuals,
                                new HtmlRenderBookmarkAnchor(
                                    segment.Run.SemanticNodeId.Value,
                                    segment.Run.BookmarkAnchorText!,
                                    x,
                                    atomicY,
                                    Math.Max(0.01D, segment.Width),
                                    Math.Max(0.01D, atomic.Height),
                                    visuals.Count,
                                    segment.Run.Source).TranslateRelativePaint(segment.Run.PaintOffsetX, segment.Run.PaintOffsetY, visuals.Count),
                                segment.Run.OwnerElement,
                                formattingContainer);
                        }
                        if (segment.Run.LinkUri != null && !AtomicRunOwnsAnchorFragment(segment.Run, atomic)) {
                            OfficeShape linkArea = OfficeShape.Rectangle(Math.Max(0.01D, segment.Width), Math.Max(0.01D, atomic.Height));
                            linkArea.FillColor = null;
                            linkArea.StrokeWidth = 0D;
                            HtmlRenderVisual linkVisual = new HtmlRenderShape(linkArea, x, atomicY, visuals.Count, segment.Run.LinkUri, segment.Run.Source);
                            if (Math.Abs(segment.Run.PaintOffsetX) > 0.0001D || Math.Abs(segment.Run.PaintOffsetY) > 0.0001D) {
                                linkVisual = linkVisual.TranslateRelativePaint(segment.Run.PaintOffsetX, segment.Run.PaintOffsetY, visuals.Count);
                            }
                            AddInlineOwnedVisual(visuals, ownedVisuals, ApplyInlineElementSemantics(linkVisual, segment.Run), segment.Run.OwnerElement, formattingContainer);
                        }
                    }
                } else if (segment.Text.Length > 0) {
                    double textLineHeight = current.HasReplacedImage || alignTextBaseline || emptyAtomicBaseline ? segment.Run.Style.LineHeight : lineHeight;
                    if (current.HasReplacedImage || alignTextBaseline || emptyAtomicBaseline) {
                        // A smaller run shares the line baseline without extending
                        // its inherited leading below the containing line box.
                        double baselineOffset = Math.Max(0D, baseline - segment.Run.Style.Font.Size);
                        textLineHeight = Math.Min(textLineHeight, Math.Max(0.01D, lineHeight - baselineOffset));
                    }
                    ResolveInlineTextVerticalPlacement(segment, current.HasReplacedImage || alignTextBaseline || emptyAtomicBaseline, lineY,
                        textLineHeight, baseline, out double textY, out double paintHeight,
                        out double paintTopOverflow);
                    RecordInlineOwnerGeometry(segment.Run, formattingContainer, x,
                        textY - paintTopOverflow, Math.Max(0.01D, segment.Width),
                        paintHeight + paintTopOverflow, inlineBounds);
                    double anchorY = lineY;
                    double anchorHeight = lineHeight;
                    if (segment.Run.LinkUri != null && !string.IsNullOrWhiteSpace(segment.Text)) {
                        if (!current.HasReplacedImage) {
                            ResolveInlineAnchorTextVerticalBounds(segment, lineY, lineHeight,
                                out anchorY, out anchorHeight);
                        }
                        RecordInlineAnchorGeometry(segment.Run, formattingContainer, x, anchorY,
                            segment.Width, anchorHeight, anchorBounds);
                    }
                    if (!segment.Run.Style.PaintVisible) {
                        cursor += rightToLeftLine ? -segment.Advance : segment.Advance;
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
                            Math.Max(0.01D, paintHeight),
                            segment.Run.Style.VectorTextDecoration
                                ? new OfficeFontInfo(segment.Run.Style.Font.FamilyName, segment.Run.Style.Font.Size,
                                    segment.Run.Style.FontDescriptor, segment.Run.Style.Font.Style & ~(OfficeFontStyle.Underline | OfficeFontStyle.Strikethrough))
                                : segment.Run.Style.Font,
                            segment.Run.Style.Color,
                            OfficeTextAlignment.Left,
                            textLineHeight,
                            textVisuals.Count,
                            segment.Run.LinkUri,
                            segment.Run.Source,
                            segment.Run.SemanticRole,
                            layoutY: lineY,
                            semanticNodeId: segment.Run.SemanticNodeId,
                            textAdvanceWidth: paintSegment.Advance,
                            bidiVisualOrderResolved: segment.BidiResolved,
                            semanticFragmentOrder: segment.Run.SemanticFragmentOrder,
                            logicalTextOrder: segment.Run.LogicalTextOrder,
                            underlineStyle: segment.Run.Style.VectorTextDecoration ? OfficeTextDecorationStyle.None : segment.Run.Style.UnderlineStyle,
                            strikethroughStyle: segment.Run.Style.VectorTextDecoration ? OfficeTextDecorationStyle.None : segment.Run.Style.StrikethroughStyle,
                            baseline: segment.Run.Style.Baseline,
                            baselineLevel: segment.Run.Style.BaselineLevel,
                            baselineScale: segment.Run.Style.BaselineScale,
                            baselineOffset: segment.Run.Style.BaselineOffset,
                            textPaintWidth: paintSegment.Width,
                            decorationColor: segment.Run.Style.DecorationColor,
                            featureSettings: segment.Run.Style.TextFeatureSettings,
                            fontPalette: segment.Run.Style.FontPalette,
                            layoutHeight: textLineHeight,
                            fontDescriptor: segment.Run.Style.FontDescriptor,
                            paintTopOverflow: paintTopOverflow,
                            linkBounds: segment.Run.LinkUri == null ? null : new HtmlRenderRectangle(
                                paintSegment.X, anchorY, Math.Max(0.01D, Math.Max(paintSegment.Width, paintSegment.Advance)), anchorHeight)));
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
                            Math.Max(0.01D, paintHeight),
                            textVisuals,
                            visuals.Count,
                            segment.Run.Source,
                            layoutY: lineY,
                            layoutHeight: textLineHeight)
                        : textVisuals[0];
                    AddInlineTextDecorations(visuals, ownedVisuals, segment.Run, formattingContainer, textVisuals, aboveText: false);
                    AddTextShadowVisuals(
                        visuals,
                        ownedVisuals,
                        segment.Run,
                        formattingContainer,
                        textVisuals);
                    AddInlineOwnedVisual(
                        visuals,
                        ownedVisuals,
                        ApplyInlineElementSemantics(textVisual.TranslateRelativePaint(segment.Run.PaintOffsetX, segment.Run.PaintOffsetY, visuals.Count), segment.Run),
                        segment.Run.OwnerElement,
                        formattingContainer);
                    AddInlineTextDecorations(visuals, ownedVisuals, segment.Run, formattingContainer, textVisuals, aboveText: true);
                }
                cursor += rightToLeftLine ? -segment.Advance : segment.Advance;
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
        double[] lineBreakOffsets = breakOffsets.ToArray();
        if (floatPlacements != null && floatPlacements.Count > 0) {
            // A line may end while a floated box still occupies the next line band.
            // That line end is not a safe page break for the float's paint.
            const double breakTolerance = 0.0001D;
            bool CrossesFloat(double offset) => floatPlacements.Any(placement =>
                offset > placement.Y + breakTolerance && offset < placement.Bottom - breakTolerance);
            breakOffsets.RemoveWhere(CrossesFloat);
            breakProgress.RemoveAll(progress => CrossesFloat(progress.Offset));
        }
        if (lineBreakGroups?.Count > 0 && flowY > 0D) {
            // Explicit float groups must not suppress the surrounding paragraph's
            // own widow/orphan group when the flow block is constructed.
            lineBreakGroups.Add(new HtmlRenderLineBreakGroup(
                lineBreakOffsets.Where(offset => offset < flowY - 0.0001D),
                paragraphStyle.Orphans, paragraphStyle.Widows, hasImplicitFinalLine: true, end: flowY));
        }
        AppendInlineAnchorFragments(visuals, ownedVisuals, anchorBounds,
            formattingContainer, isInlineContinuation);
        paintCapture?.Capture(visuals, ownedVisuals, inlineBounds);
        IReadOnlyList<HtmlRenderVisual> painted = paintCapture == null
            ? ComposeInlinePositionedVisuals(visuals, ownedVisuals, inlineBounds, formattingContainer, isInlineContinuation)
            : visuals.Concat(ownedVisuals.Values.SelectMany(items => items)).ToArray();
        // Interrupted inline chunks have their own float context. Their parent
        // cannot recover its row boundaries from the shared block context.
        IEnumerable<double> pageBreakOffsets = localFloatPageHeight > 0D && hasFloats
            ? CollectSafeFloatBreaks(breakOffsets, floatPlacements!, CollectAtomicParallelVisualRanges(painted), localFloatPageHeight)
            : breakOffsets;
        return new HtmlInlineLayout(
            painted,
            height,
            pageBreakOffsets,
            runningStringAssignments.OrderBy(assignment => assignment.OrderOffset),
            breakProgress,
            supportsContinuationReflow,
            flowY,
            lineBreakOffsets,
            (floatPlacements?.Select(placement => new HtmlFloatExclusion(
                placement.X, placement.Y, placement.Width, placement.Height, placement.Run.FloatSide))
                ?? Enumerable.Empty<HtmlFloatExclusion>()).Concat(flowExclusions ?? Array.Empty<HtmlFloatExclusion>()),
            forcedBreaks: forcedBreaks,
            lineBreakGroups: lineBreakGroups,
            continuationGroups: continuationGroups,
            trailingGroups: trailingGroups,
            pagedPaintExtent: hasFloats ? floatPlacements!.Max(placement => placement.Bottom) : height,
            hasInFlowLineBoxes: lines.Any(line => line.HasFlowContent || line.HasForcedBreak));
    }

}
