using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private static bool IsSpecializedTableCell(IElement element) => IsReplacedImageElement(element)
        || IsFormControlElement(element.LocalName) && !UsesButtonChildLayout(element);

    private HtmlInlineLayout LayoutTableCellContent(
        TableFormattingCell cell,
        double contentWidth,
        HtmlRenderBoxStyle style,
        int depth,
        bool paintSeparateBorders) {
        bool previous = _tableCellHeightBasisActive;
        _tableCellHeightBasisActive = true;
        try { return LayoutTableCellContentCore(cell, contentWidth, style, depth, paintSeparateBorders); }
        finally { _tableCellHeightBasisActive = previous; }
    }

    private HtmlInlineLayout LayoutTableCellContentCore(
        TableFormattingCell cell,
        double contentWidth,
        HtmlRenderBoxStyle style,
        int depth,
        bool paintSeparateBorders) {
        if (cell.Nodes == null && IsSpecializedTableCell(cell.Element)) {
            HtmlRenderBoxStyle boxStyle = style.Clone();
            // Keep border layout insets here; the table retains the original
            // edges for conflict resolution and owns their collapsed paint.
            if (!paintSeparateBorders) boxStyle.Borders = boxStyle.Borders.WithUniformColor(OfficeIMO.Drawing.OfficeColor.Transparent);
            boxStyle.MarginLeft = boxStyle.MarginRight = boxStyle.MarginTop = boxStyle.MarginBottom = 0D;
            boxStyle.ExplicitWidth = contentWidth + (boxStyle.BorderBox ? boxStyle.HorizontalInsets : 0D);
            boxStyle.ExplicitWidthUsesPercentage = false;
            HtmlRenderFlowBlock box = IsReplacedImageElement(cell.Element)
                ? LayoutImage(cell.Element, contentWidth + style.HorizontalInsets, boxStyle)
                : LayoutFormControl(cell.Element, contentWidth + style.HorizontalInsets, boxStyle);
            // The specialized renderer owns its complete box. Table content is
            // translated from the content origin, so remove its inset once here.
            double x = -style.BorderLeftWidth - style.PaddingLeft;
            double y = -style.BorderTopWidth - style.PaddingTop;
            return new HtmlInlineLayout(box.Visuals.Select(visual => visual.Translate(x, y, visual.PaintOrder)).ToArray(),
                Math.Max(0D, box.Height - style.VerticalInsets), Array.Empty<double>(), box.RunningStringAssignments);
        }
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
