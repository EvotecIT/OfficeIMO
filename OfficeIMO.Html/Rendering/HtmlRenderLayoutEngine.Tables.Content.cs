using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private HtmlInlineLayout LayoutTableCellContent(
        IElement cell,
        double contentWidth,
        HtmlRenderBoxStyle style,
        int depth) {
        if (!HasBlockChildren(cell, contentWidth, style)) {
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
        return new HtmlInlineLayout(visuals, height, breakOffsets);
    }
}
