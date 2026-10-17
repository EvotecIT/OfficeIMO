using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private HtmlInlineLayout LayoutTableCellContent(
        TableFormattingCell cell,
        double contentWidth,
        HtmlRenderBoxStyle style,
        int depth) {
        if (cell.Nodes == null && !HasBlockChildren(cell.Element, contentWidth, style, depth)) {
            return LayoutInlineNodes(cell.Element.ChildNodes, contentWidth, style, depth, null, cell.Element);
        }

        IReadOnlyList<HtmlRenderFlowBlock> blocks = cell.Nodes == null
            ? BuildChildBlocks(cell.Element, contentWidth, style, depth)
            : BuildChildBlocks(cell.Element, cell.Nodes, contentWidth, style, depth, includeGeneratedBefore: false, includeGeneratedAfter: false);
        var visuals = new List<HtmlRenderVisual>();
        var paintLayers = new List<FlowPaintLayer>();
        var breakOffsets = new SortedSet<double>();
        double height = 0D;
        foreach (HtmlRenderFlowBlock block in blocks) {
            double blockStart = height;
            paintLayers.Add(new FlowPaintLayer(block, 0D, blockStart, paintLayers.Count));
            foreach (double offset in block.BreakOffsets) {
                double translated = blockStart + offset;
                if (translated > 0D) breakOffsets.Add(translated);
            }
            height += block.Height;
        }

        AppendFlowPaintLayers(visuals, paintLayers);
        return new HtmlInlineLayout(
            visuals,
            height,
            breakOffsets,
            paintLayers.SelectMany(layer =>
                layer.Block.RunningStringAssignments.Select(assignment => assignment.Translate(layer.Y))));
    }
}
