using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderFlowBlock {
    // Re-normalizing these owned lists is idempotent only for an unchanged finite extent.
    // Arbitrary and geometry-changing inputs retain the general constructor.
    private bool CanShareNormalizedFacts => Width >= 0D && !double.IsInfinity(Width)
        && Height >= 0D && !double.IsInfinity(Height)
        && PagedPaintExtent >= 0D && !double.IsInfinity(PagedPaintExtent)
        && !double.IsNaN(UnclampedHeight) && !double.IsInfinity(UnclampedHeight);

    /// <summary>Copies unchanged geometry and owned normalized facts; replacement paint is still snapshotted eagerly.</summary>
    private HtmlRenderFlowBlock(HtmlRenderFlowBlock source, IEnumerable<HtmlRenderVisual> visuals,
        double layoutViewportWidth, double layoutViewportHeight) {
        Width = source.Width;
        Height = source.Height;
        PagedPaintExtent = source.PagedPaintExtent;
        UnclampedHeight = source.UnclampedHeight;
        Visuals = new List<HtmlRenderVisual>(visuals);
        BreakBefore = source.BreakBefore;
        BreakAfter = source.BreakAfter;
        AvoidBreakInside = source.AvoidBreakInside;
        AvoidBreakInsideScope = AvoidBreakInside ? HtmlFragmentationScope.All : HtmlFragmentationScope.None;
        Source = source.Source;
        BreakOffsets = source.BreakOffsets;
        SafeBreakRanges = source.SafeBreakRanges;
        AvoidBreakRanges = source.AvoidBreakRanges;
        ForcedBreaks = source.ForcedBreaks;
        LineBreakGroups = source.LineBreakGroups;
        ContinuationGroups = source.ContinuationGroups;
        TrailingGroups = source.TrailingGroups;
        PageName = source.PageName;
        StackingZIndex = source.StackingZIndex;
        StackingSourceOrder = source.StackingSourceOrder;
        HasCollapsibleMargins = source.HasCollapsibleMargins;
        CollapsibleMarginTop = source.CollapsibleMarginTop;
        CollapsibleMarginBottom = source.CollapsibleMarginBottom;
        CollapsibleMarginTopGroup = source.CollapsibleMarginTopGroup;
        CollapsibleMarginBottomGroup = source.CollapsibleMarginBottomGroup;
        OwnerElement = source.OwnerElement;
        CollapsesThrough = source.CollapsesThrough;
        RunningStringAssignments = source.RunningStringAssignments;
        InlineBreakProgress = source.InlineBreakProgress;
        InlineContinuationStart = source.InlineContinuationStart;
        SupportsInlineContinuationReflow = source.SupportsInlineContinuationReflow;
        LayoutViewportWidth = layoutViewportWidth;
        LayoutViewportHeight = layoutViewportHeight;
        LeadingFlowAdjustment = source.LeadingFlowAdjustment;
        // CopyLayoutFacts still applies the derived flags, overflow projection and height clamp.
        // KnownNoInFlowLineBoxes deliberately retains the general constructor's false default.
    }

    private HtmlRenderFlowBlock(HtmlRenderFlowBlock source, double top, double bottom, IElement ownerElement,
        bool collapsesThrough, HtmlCollapsedMargin? topGroup, HtmlCollapsedMargin? bottomGroup)
        : this(source, source.Visuals, source.LayoutViewportWidth, source.LayoutViewportHeight) {
        HasCollapsibleMargins = true;
        CollapsibleMarginTop = top;
        CollapsibleMarginBottom = bottom;
        CollapsibleMarginTopGroup = topGroup ?? new HtmlCollapsedMargin(top);
        CollapsibleMarginBottomGroup = bottomGroup ?? new HtmlCollapsedMargin(bottom);
        OwnerElement = ownerElement;
        CollapsesThrough = collapsesThrough;
        List<HtmlRenderLineBreakGroup>? changedGroups = null;
        for (int index = 0; index < LineBreakGroups.Count; index++) {
            HtmlRenderLineBreakGroup group = LineBreakGroups[index];
            if (!(bottom > 0.0001D && Math.Abs(group.End - (Height - bottom)) <= 0.0001D)) continue;
            changedGroups ??= new List<HtmlRenderLineBreakGroup>(LineBreakGroups);
            changedGroups[index] = group.WithFinalLineMarginBreak();
        }
        if (changedGroups != null) LineBreakGroups = changedGroups.AsReadOnly();
    }
}
