namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    /// <summary>Keeps a clipped viewport independent of its unfragmented internal source.</summary>
    private HtmlRenderFlowBlock ProjectAtomicOrdinaryBlock(HtmlRenderFlowBlock block, PagedFloatBoundary boundary,
        double viewportStart, double viewportEnd) {
        ConstrainedSourceProjection? context = _constrainedSourceProjection;
        if (context == null || context.FailedReason != null) return block;
        if (viewportEnd > context.PageHeight + 0.0001D) {
            context.FailedReason = "oversized-clipped-source-owner";
            return block;
        }
        double available = boundary.RemainingHeight;
        if (available <= 0.0001D)
            available += (Math.Floor(-available / context.PageHeight) + 1D) * context.PageHeight;
        double gap = viewportEnd > available + 0.0001D ? available : 0D;
        if (gap > 0.0001D) {
            context.Changed = true;
            context.NeedsProjection = true;
        }
        HtmlRenderFlowBlock positioned = gap > 0.0001D ? block.AdjustLeadingFlowSpace(-gap) : block;
        CheckCancellation();
        ChargeLayoutOperations(1L + positioned.Visuals.Count + positioned.RunningStringAssignments.Count
            + positioned.InlineBreakProgress.Count + positioned.ForcedBreaks.Count,
            "clipped source projection isolation");
        // Internal line/keep opportunities describe hidden, unfragmented content.
        // The shared outer cursor owns only this viewport and its entry spacing.
        return new HtmlRenderFlowBlock(positioned.Width, positioned.Height, positioned.Visuals,
            positioned.BreakBefore, positioned.BreakAfter, false, positioned.Source,
            breakOffsets: new[] { gap, positioned.Height },
            pageName: positioned.PageName,
            stackingZIndex: positioned.StackingZIndex,
            stackingSourceOrder: positioned.StackingSourceOrder,
            hasCollapsibleMargins: positioned.HasCollapsibleMargins,
            collapsibleMarginTop: positioned.CollapsibleMarginTop,
            collapsibleMarginBottom: positioned.CollapsibleMarginBottom,
            ownerElement: positioned.OwnerElement,
            collapsesThrough: positioned.CollapsesThrough,
            unclampedHeight: positioned.UnclampedHeight,
            runningStringAssignments: positioned.RunningStringAssignments,
            inlineBreakProgress: positioned.InlineBreakProgress.Where(item => item.Offset <= gap + 0.0001D
                || item.Offset >= positioned.Height - 0.0001D),
            inlineContinuationStart: positioned.InlineContinuationStart,
            supportsInlineContinuationReflow: false,
            forcedBreaks: positioned.ForcedBreaks.Where(item => item.Offset <= gap + 0.0001D
                || item.Offset >= positioned.Height - 0.0001D),
            layoutViewportWidth: positioned.LayoutViewportWidth,
            layoutViewportHeight: positioned.LayoutViewportHeight,
            leadingFlowAdjustment: positioned.LeadingFlowAdjustment,
            collapsibleMarginTopGroup: positioned.CollapsibleMarginTopGroup,
            collapsibleMarginBottomGroup: positioned.CollapsibleMarginBottomGroup,
            avoidBreakRanges: new[] { new HtmlRenderAvoidBreakRange(Math.Max(0D, viewportStart + gap), viewportEnd + gap,
                Scope: HtmlFragmentationScope.Page | positioned.AvoidBreakInsideScope) },
            pagedPaintExtent: positioned.PagedPaintExtent,
            safeBreakRanges: gap > 0.0001D ? new[] { new HtmlRenderBreakRange(0D, gap) } : null) {
                RetainsZeroHeightFlowCarrier = positioned.RetainsZeroHeightFlowCarrier,
                KnownNoInFlowLineBoxes = positioned.KnownNoInFlowLineBoxes,
                HasZeroHeightOrdinaryContent = positioned.HasZeroHeightOrdinaryContent,
                PagedSourceEntryGap = gap,
                SourceEntryNode = positioned.SourceEntryNode,
                UnfragmentedHeight = positioned.UnfragmentedHeight,
                ColumnOverflowContent = positioned.ColumnOverflowContent
            };
    }
}
