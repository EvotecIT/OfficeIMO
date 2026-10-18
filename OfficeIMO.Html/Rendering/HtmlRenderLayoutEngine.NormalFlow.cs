using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private static bool HasSignedFlowAllocation(IReadOnlyList<HtmlRenderFlowBlock> blocks) =>
        blocks.Any(block => Math.Abs(block.NormalFlowAdvance - block.Height) > 0.0001D);

    /// <summary>Imports explicit paint overhang without turning a flow carrier's allocation into painted space.</summary>
    private static double ResolveNormalFlowPaintExtent(HtmlRenderFlowBlock block) =>
        block.PagedPaintExtent > block.Height + 0.0001D
            ? block.PagedPaintExtent : Math.Max(0D, block.NormalFlowAdvance);

    /// <summary>
    /// Shares one source and paint surface when overlapping siblings cannot be
    /// represented by consecutive nonnegative fragment allocations. Children
    /// keep their original visuals and fragmentation facts at signed positions;
    /// only the enclosing fragmentable extent is nonnegative.
    /// </summary>
    private HtmlRenderFlowBlock BuildNormalFlowSurface(
        IReadOnlyList<HtmlRenderFlowBlock> blocks, double width, IElement owner,
        bool retainsRootReflow = false) {
        var visuals = new List<HtmlRenderVisual>();
        var layers = new List<FlowPaintLayer>();
        var offsets = new SortedSet<double>();
        var safe = new List<HtmlRenderBreakRange>();
        var keeps = new List<HtmlRenderAvoidBreakRange>();
        var lines = new List<HtmlRenderLineBreakGroup>();
        var continuations = new List<HtmlRenderContinuationGroup>();
        var trailing = new List<HtmlRenderTrailingGroup>();
        var assignments = new List<HtmlCssRunningStringAssignment>();
        var progress = new List<HtmlInlineBreakProgress>();
        var forced = new List<HtmlRenderForcedBreak>();
        object[] boundaries = CreateFlowBoundaries(blocks);
        string? pageName = blocks.FirstOrDefault()?.PageName;
        double advance = 0D;
        double extent = 0D;
        for (int index = 0; index < blocks.Count; index++) {
            CheckCancellation();
            HtmlRenderFlowBlock block = blocks[index];
            ChargeLayoutOperations(1L + block.Visuals.Count + block.BreakOffsets.Count
                + block.SafeBreakRanges.Count + block.AvoidBreakRanges.Count
                + block.LineBreakGroups.Count + block.ContinuationGroups.Count
                + block.TrailingGroups.Count + block.RunningStringAssignments.Count
                + block.InlineBreakProgress.Count + block.ForcedBreaks.Count,
                "overlapping normal-flow surface");
            if (index > 0 && !string.Equals(pageName, block.PageName, StringComparison.Ordinal))
                forced.Add(new HtmlRenderForcedBreak(advance, HtmlPageBreakTarget.Page,
                    block.PageName, changesPageName: true));
            pageName = block.PageName;
            AppendChildForcedBreaks(blocks, index, advance, forced, boundaries);
            AppendKeepWithNextRange(blocks, index, advance, keeps);
            if (block.AvoidBreakInsideScope != HtmlFragmentationScope.None)
                keeps.Add(new HtmlRenderAvoidBreakRange(advance + block.PagedSourceEntryGap, advance + block.Height,
                    Scope: block.AvoidBreakInsideScope));
            layers.Add(new FlowPaintLayer(block, 0D, advance, layers.Count));
            foreach (double offset in block.BreakOffsets) offsets.Add(advance + offset);
            safe.AddRange(block.SafeBreakRanges.Select(range => range.Translate(advance)));
            keeps.AddRange(block.AvoidBreakRanges.Select(range => range.Translate(advance)));
            lines.AddRange(block.LineBreakGroups.Select(group => group.Translate(advance)));
            continuations.AddRange(block.ContinuationGroups.Select(group => group.Translate(0D, advance)));
            trailing.AddRange(block.TrailingGroups.Select(group => group.Translate(0D, advance)));
            assignments.AddRange(block.RunningStringAssignments.Select(item => item.Translate(advance)));
            progress.AddRange(block.InlineBreakProgress.Select(item => new HtmlInlineBreakProgress(
                advance + item.Offset, item.LogicalCharacters, item.OwnerElement,
                item.IsBlockEntry, item.PageStartDiscardableMargin, item.IsBlockExit, item.IsFlexGap)));
            extent = Math.Max(extent, advance + ResolveNormalFlowPaintExtent(block));
            advance += block.NormalFlowAdvance;
            offsets.Add(advance);
        }
        AppendFlowPaintLayers(visuals, layers);
        double height = Math.Max(0D, advance);
        extent = Math.Max(height, extent);
        bool hasColumnOverflow = _columnContentDepth > 0
            && layers.Any(layer => layer.Block.ColumnOverflowContent != null);
        double columnOverflowExtent = hasColumnOverflow
            ? AppendColumnChildOverflowFacts(layers, 0D, extent, offsets, safe, keeps, forced) : extent;
        var surface = new HtmlRenderFlowBlock(width, height, visuals,
            HtmlPageBreakTarget.None, HtmlPageBreakTarget.None, false,
            HtmlRenderStyleResolver.DescribeSource(owner), offsets,
            lineBreakGroups: lines, continuationGroups: continuations, trailingGroups: trailing,
            pageName: blocks.FirstOrDefault()?.PageName, unclampedHeight: advance,
            ownerElement: retainsRootReflow ? owner : null,
            runningStringAssignments: assignments, inlineBreakProgress: progress,
            supportsInlineContinuationReflow: blocks.Any(child => child.SupportsInlineContinuationReflow),
            layoutViewportWidth: ActiveSurfaceWidth, layoutViewportHeight: _activePageGeometry.Height,
            forcedBreaks: forced, avoidBreakRanges: keeps, pagedPaintExtent: columnOverflowExtent,
            safeBreakRanges: ConstrainParallelPaintBreakRanges(
                HtmlRenderBreakRange.Normalize(safe, columnOverflowExtent), visuals));
        if (hasColumnOverflow && columnOverflowExtent > height + 0.0001D) {
            HtmlRenderFlowBlock overflow = surface.ForPagination();
            surface = surface.WithVisuals(surface.Visuals, pagedPaintExtent: extent);
            surface.ColumnOverflowContent = overflow;
        }
        return surface;
    }
}
