namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    /// <summary>Maps one source gap without changing unrelated paint or source owners.</summary>
    private HtmlRenderFlowBlock InsertPagedSourceBreakGap(HtmlRenderFlowBlock block, double cut, double gap,
        bool absorbInStretchTail, double? flowAdvanceDelta = null, PagedBackgroundWindow? backgroundWindow = null,
        object? projectedForcedBoundary = null) {
        CheckCancellation();
        ChargeLayoutOperations(1L + block.Visuals.Count + block.BreakOffsets.Count
            + block.SafeBreakRanges.Count + block.LineBreakGroups.Count + block.AvoidBreakRanges.Count
            + block.ContinuationGroups.Count + block.TrailingGroups.Count + block.ForcedBreaks.Count
            + block.InlineBreakProgress.Count + block.RunningStringAssignments.Count, "paged source gap mapping");
        double advanceDelta = absorbInStretchTail ? 0D : flowAdvanceDelta ?? gap;
        IReadOnlyList<HtmlRenderVisual> before = SliceVisuals(block.Visuals, 0D, cut, backgroundWindow: backgroundWindow);
        IReadOnlyList<HtmlRenderVisual> after = SliceVisuals(block.Visuals, cut, block.PagedPaintExtent, includeEndAnchor: true);
        List<HtmlRenderVisual> visuals = before
            .Concat(after.Select((visual, index) => visual.Translate(0D, cut + gap, before.Count + index)))
            .ToList();
        double height = Math.Max(0D, block.Height + advanceDelta);
        double extent = absorbInStretchTail ? height : Math.Max(height, block.PagedPaintExtent + gap);
        if (absorbInStretchTail) visuals = SliceVisuals(visuals, 0D, height, includeEndAnchor: true).ToList();
        double Shift(double offset) => offset >= cut - 0.0001D ? offset + gap : offset;
        HtmlRenderForcedBreak MapForcedBreak(HtmlRenderForcedBreak item) {
            if (projectedForcedBoundary != null && Math.Abs(item.Offset - cut) <= 0.0001D) {
                // A prior empty boundary already owns the preceding page edge.
                // Later distinct boundaries at this offset still move with their source.
                if (item.Projected) return item;
                if (ReferenceEquals(item.Boundary, projectedForcedBoundary)) return item.ProjectAt(item.Offset + gap);
            }
            return item.Translate(Shift(item.Offset) - item.Offset);
        }

        return new HtmlRenderFlowBlock(
            block.Width,
            height,
            visuals,
            block.BreakBefore,
            block.BreakAfter,
            block.AvoidBreakInside,
            block.Source,
            block.BreakOffsets.Select(Shift).Concat(new[] { cut, cut + gap }),
            lineBreakGroups: block.LineBreakGroups.Select(group => new HtmlRenderLineBreakGroup(
                group.Offsets.Select(Shift),
                group.Orphans,
                group.Widows,
                group.HasImplicitFinalLine,
                group.Start >= cut - 0.0001D ? group.Start + gap : group.Start,
                group.End > cut + 0.0001D ? group.End + gap : group.End,
                group.CheckInteriorBreaks,
                group.FinalLineMarginBreak)),
            continuationGroups: block.ContinuationGroups.Select(group => new HtmlRenderContinuationGroup(
                Shift(group.StartsAfter), Shift(group.EndsAt), group.Height, group.Visuals)),
            trailingGroups: block.TrailingGroups.Select(group => new HtmlRenderTrailingGroup(
                Shift(group.StartsAt), Shift(group.ContentEndsAt), Shift(group.SourceEndsAt), group.Height, group.Visuals)),
            forcedBreaks: block.ForcedBreaks.Select(MapForcedBreak),
            pageName: block.PageName,
            stackingZIndex: block.StackingZIndex,
            stackingSourceOrder: block.StackingSourceOrder,
            hasCollapsibleMargins: block.HasCollapsibleMargins,
            collapsibleMarginTop: block.CollapsibleMarginTop,
            collapsibleMarginBottom: block.CollapsibleMarginBottom,
            ownerElement: block.OwnerElement,
            collapsesThrough: block.CollapsesThrough,
            unclampedHeight: block.UnclampedHeight + advanceDelta,
            runningStringAssignments: block.RunningStringAssignments.Select(assignment =>
                assignment.Offset >= cut - 0.0001D ? assignment.Translate(gap) : assignment),
            inlineBreakProgress: block.InlineBreakProgress.Select(progress => new HtmlInlineBreakProgress(
                Shift(progress.Offset), progress.LogicalCharacters, progress.OwnerElement,
                progress.IsBlockEntry, progress.PageStartDiscardableMargin, progress.IsBlockExit, progress.IsFlexGap)),
            inlineContinuationStart: block.InlineContinuationStart,
            supportsInlineContinuationReflow: block.SupportsInlineContinuationReflow,
            layoutViewportWidth: block.LayoutViewportWidth,
            layoutViewportHeight: block.LayoutViewportHeight,
            leadingFlowAdjustment: block.LeadingFlowAdjustment,
            collapsibleMarginTopGroup: block.CollapsibleMarginTopGroup,
            collapsibleMarginBottomGroup: block.CollapsibleMarginBottomGroup,
            avoidBreakRanges: block.AvoidBreakRanges.Select(range => new HtmlRenderAvoidBreakRange(
                range.Start >= cut - 0.0001D ? range.Start + gap : range.Start,
                range.End > cut + 0.0001D ? range.End + gap : range.End,
                range.Soft, range.Scope)),
            pagedPaintExtent: extent,
            safeBreakRanges: SplitPagedSourceBreakRangesAtGap(block.SafeBreakRanges, cut, gap, extent, backgroundWindow.HasValue)) {
                AvoidBreakInsideScope = block.AvoidBreakInsideScope,
                RetainsZeroHeightFlowCarrier = block.RetainsZeroHeightFlowCarrier,
                KnownNoInFlowLineBoxes = block.KnownNoInFlowLineBoxes,
                HasZeroHeightOrdinaryContent = block.HasZeroHeightOrdinaryContent,
                PagedSourceEntryGap = block.PagedSourceEntryGap,
                UnfragmentedHeight = block.UnfragmentedHeight,
                ColumnOverflowContent = block.ColumnOverflowContent == null ? null
                    : InsertPagedSourceBreakGap(block.ColumnOverflowContent, cut, gap, absorbInStretchTail, flowAdvanceDelta, backgroundWindow, projectedForcedBoundary)
            };
    }

    /// <summary>Maps each side of the retained source independently; inserted spacing is not a source interval.</summary>
    private IReadOnlyList<HtmlRenderBreakRange> SplitPagedSourceBreakRangesAtGap(
        IReadOnlyList<HtmlRenderBreakRange> ranges, double cut, double gap, double extent, bool includeGap) {
        var result = new List<HtmlRenderBreakRange>();
        if (includeGap) result.Add(new HtmlRenderBreakRange(cut, cut + gap));
        foreach (HtmlRenderBreakRange range in ranges) {
            CheckCancellation();
            ChargeLayoutOperation("paged source break interval gap mapping");
            if (range.Start <= cut && range.End >= 0D)
                result.Add(new HtmlRenderBreakRange(range.Start, Math.Min(cut, range.End)));
            if (range.End >= cut)
                result.Add(new HtmlRenderBreakRange(Math.Max(cut, range.Start) + gap, range.End + gap));
        }
        return HtmlRenderBreakRange.Normalize(result, extent);
    }
}
