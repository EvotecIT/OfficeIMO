using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private HtmlInlineLayout LayoutTableCellContent(
        IElement cell,
        double contentWidth,
        HtmlRenderBoxStyle style,
        int depth) {
        // A cell's authored height is a minimum for the row, not a definite
        // containing height for its children. Orthogonal inline blocks must
        // be able to expand the row before their own height limit is applied.
        if (style.ExplicitHeight.HasValue) {
            style = style.Clone();
            style.ExplicitHeight = null;
        }
        if (!HasBlockChildren(cell, contentWidth, style, depth)) {
            return LayoutInlineNodes(cell.ChildNodes, contentWidth, style, depth, null, cell);
        }

        IReadOnlyList<HtmlRenderFlowBlock> blocks = BuildChildBlocks(cell, contentWidth, style, depth);
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
