using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private HtmlInlineLayout LayoutTableCellContent(
        TableFormattingCell cell,
        double contentWidth,
        HtmlRenderBoxStyle style,
        int depth) => LayoutTableBoxContent(cell.Element, contentWidth, style, depth, cell.Nodes,
            cell.Nodes == null || cell.GeneratedBefore, cell.Nodes == null || cell.GeneratedAfter);

    /// <summary>Lays out cell or caption content through the same inline and block formatting paths.</summary>
    private HtmlInlineLayout LayoutTableBoxContent(IElement element, double contentWidth, HtmlRenderBoxStyle style,
        int depth, IReadOnlyList<INode>? sourceNodes = null, bool includeGeneratedBefore = true, bool includeGeneratedAfter = true) {
        using var pageBoundaries = new DescendantPageBoundaryScope(this, retainsPageBoundaries: false);
        if (sourceNodes == null && !HasBlockChildren(element, contentWidth, style, depth)) {
            return LayoutInlineNodes(element.ChildNodes, contentWidth, style, depth, null, element);
        }
        if (sourceNodes != null && (includeGeneratedBefore || includeGeneratedAfter)
            && !HasBlockChildren(element, contentWidth, style, depth, sourceNodes, includeGeneratedBefore, includeGeneratedAfter)) {
            return LayoutInlineNodes(sourceNodes, contentWidth, style, depth, null, element,
                includeGeneratedBefore: includeGeneratedBefore, includeGeneratedAfter: includeGeneratedAfter);
        }

        IReadOnlyList<HtmlRenderFlowBlock> blocks = sourceNodes == null
            ? BuildChildBlocks(element, contentWidth, style, depth)
            : BuildChildBlocks(element, sourceNodes, contentWidth, style, depth,
                includeGeneratedBefore: includeGeneratedBefore, includeGeneratedAfter: includeGeneratedAfter);
        var visuals = new List<HtmlRenderVisual>();
        var paintLayers = new List<FlowPaintLayer>();
        var breakOffsets = new SortedSet<double>();
        var safeBreakRanges = new List<HtmlRenderBreakRange>();
        var avoidBreakRanges = new List<HtmlRenderAvoidBreakRange>();
        var lineBreakGroups = new List<HtmlRenderLineBreakGroup>();
        double height = 0D;
        double paintExtent = 0D;
        for (int index = 0; index < blocks.Count; index++) {
            CheckCancellation();
            ChargeLayoutOperation("table cell fragmentation metadata");
            HtmlRenderFlowBlock block = blocks[index];
            double blockStart = height;
            AppendKeepWithNextRange(blocks, index, blockStart, avoidBreakRanges);
            paintLayers.Add(new FlowPaintLayer(block, 0D, blockStart, paintLayers.Count));
            foreach (double offset in block.BreakOffsets) {
                double translated = blockStart + offset;
                if (translated > 0D) breakOffsets.Add(translated);
            }
            safeBreakRanges.AddRange(block.SafeBreakRanges.Select(range => range.Translate(blockStart)));
            paintExtent = Math.Max(paintExtent, blockStart + ResolveNormalFlowPaintExtent(block));
            height += block.NormalFlowAdvance;
            if (block.AvoidBreakInsideScope != HtmlFragmentationScope.None)
                avoidBreakRanges.Add(new HtmlRenderAvoidBreakRange(blockStart, height, Scope: block.AvoidBreakInsideScope));
            foreach (HtmlRenderAvoidBreakRange range in block.AvoidBreakRanges) {
                ChargeLayoutOperation("table cell keep ranges");
                avoidBreakRanges.Add(range.Translate(blockStart));
            }
            foreach (HtmlRenderLineBreakGroup group in block.LineBreakGroups) {
                ChargeLayoutOperation("table cell line groups");
                lineBreakGroups.Add(group.Translate(blockStart));
            }
        }

        AppendFlowPaintLayers(visuals, paintLayers);
        return new HtmlInlineLayout(
            visuals,
            height,
            breakOffsets,
            paintLayers.SelectMany(layer =>
                layer.Block.RunningStringAssignments.Select(assignment => assignment.Translate(layer.Y))),
            lineBreakGroups: lineBreakGroups,
            avoidBreakRanges: avoidBreakRanges,
            pagedPaintExtent: paintExtent,
            safeBreakRanges: safeBreakRanges);
    }
}
