using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private TableCaptionLayout? LayoutTableCaption(
        IElement? element,
        double tableWidth,
        HtmlRenderBoxStyle tableStyle,
        int depth) {
        if (element == null) return null;

        HtmlRenderBoxStyle style = _styleResolver.Resolve(element, tableWidth, tableStyle);
        _layoutStyles[element] = style.Clone();
        if (tableStyle.UnsupportedCaptionSide.Length == 0) ReportUnsupportedTableValues(element, style);
        if (style.Display == "none") return null;

        double availableWidth = Math.Max(1D, tableWidth - style.MarginLeft - style.MarginRight);
        double boxWidth = ResolveBoxWidth(availableWidth, style);
        double contentWidth = Math.Max(1D, boxWidth - style.HorizontalInsets);
        HtmlInlineLayout inline = LayoutInlineNodes(element.ChildNodes, contentWidth, style, depth + 1, null, element);
        // A caption participates in flow through actual line boxes or authored
        // box dimensions. An empty caption does not create an implicit strut.
        double boxHeight = ResolveBoxHeight(inline.Height, boxWidth, style, allowZeroHeight: true);
        var visuals = new List<HtmlRenderVisual>();
        AddBoxPaint(visuals, style, style.MarginLeft, style.MarginTop, boxWidth, boxHeight, element);
        double contentX = style.MarginLeft + style.BorderLeftWidth + style.PaddingLeft;
        double contentY = style.MarginTop + style.BorderTopWidth + style.PaddingTop;
        foreach (HtmlRenderVisual visual in inline.Visuals) visuals.Add(visual.Translate(contentX, contentY, visuals.Count));
        AddBoxOutlinePaint(visuals, style, style.MarginLeft, style.MarginTop, boxWidth, boxHeight, element);
        IReadOnlyList<HtmlRenderVisual> semanticVisuals = new[] {
            new HtmlRenderSemanticGroup(
                HtmlRenderSemanticGroupRole.Caption,
                style.MarginLeft,
                style.MarginTop,
                boxWidth,
                Math.Max(0.01D, boxHeight),
                visuals,
                0,
                HtmlRenderStyleResolver.DescribeSource(element))
        };
        return new TableCaptionLayout(style.CaptionSide, style.MarginTop + boxHeight + style.MarginBottom, semanticVisuals);
    }

    private static void AppendTableCaption(
        ICollection<HtmlRenderVisual> target,
        TableCaptionLayout caption,
        double x,
        double y) {
        foreach (HtmlRenderVisual visual in caption.Visuals) target.Add(visual.Translate(x, y, target.Count));
    }
}
