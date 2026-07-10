using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private bool TryAddFlexNode(
        INode node,
        double containingWidth,
        HtmlRenderBoxStyle parentStyle,
        int depth,
        ref int sourceIndex,
        ICollection<FlexItem> items) {
        if (node is IText text) {
            if (string.IsNullOrWhiteSpace(text.Data)) return true;
            string source = HtmlRenderStyleResolver.DescribeSource(text.ParentElement ?? throw new InvalidOperationException("A flex text node has no parent element.")) + "::anonymous-flex-item";
            IElement owner = text.ParentElement!;
            items.Add(new FlexItem(text.Data, owner, source, ResolveFlexItemLink(owner), CreateAnonymousFlexStyle(parentStyle), sourceIndex++, paintAnonymousBox: false));
            return true;
        }

        if (!(node is IElement element) || ShouldSkipElement(element)) return true;
        EnsureDepth(depth, element);
        HtmlRenderBoxStyle style = _styleResolver.Resolve(element, containingWidth, parentStyle);
        if (style.Display == "none") return true;
        if (style.Display == "contents") {
            AddGeneratedFlexItem(element, HtmlPseudoElementKind.Before, containingWidth, style, ref sourceIndex, items);
            foreach (INode child in element.ChildNodes) {
                if (!TryAddFlexNode(child, containingWidth, style, depth + 1, ref sourceIndex, items)) return false;
            }
            AddGeneratedFlexItem(element, HtmlPseudoElementKind.After, containingWidth, style, ref sourceIndex, items);
            return true;
        }

        if (style.Position == "absolute" || style.Position == "fixed") {
            RegisterOutOfFlowElement(element.ParentElement ?? element, element, style, parentStyle, depth);
            return true;
        }
        if (style.Position != "static" && style.Position != "relative" && style.Position != "sticky") return false;
        items.Add(new FlexItem(element, BlockifyFlexItemStyle(style), sourceIndex++));
        return true;
    }

    private void AddGeneratedFlexItem(
        IElement element,
        HtmlPseudoElementKind kind,
        double containingWidth,
        HtmlRenderBoxStyle parentStyle,
        ref int sourceIndex,
        ICollection<FlexItem> items) {
        if (!_generatedContent.TryGet(element, kind, out string content)
            || !_styleResolver.TryResolvePseudo(element, kind, containingWidth, parentStyle, out HtmlRenderBoxStyle style)
            || style.Display == "none") {
            return;
        }

        string source = DescribePseudoSource(element, kind);
        ReportUnsupportedGeneratedLayout(style, source);
        items.Add(new FlexItem(
            content,
            element,
            source,
            ResolveFlexItemLink(element),
            BlockifyFlexItemStyle(style),
            sourceIndex++,
            paintAnonymousBox: true));
    }

    private HtmlRenderFlowBlock LayoutFlexItem(FlexItem item, double containingWidth, HtmlRenderBoxStyle parentStyle, int depth) {
        if (item.Element != null) return LayoutElement(item.Element, containingWidth, item.Style, parentStyle, depth);
        return LayoutAnonymousFlexItem(item, containingWidth, parentStyle);
    }

    private HtmlRenderFlowBlock LayoutAnonymousFlexItem(FlexItem item, double containingWidth, HtmlRenderBoxStyle parentStyle) {
        HtmlRenderBoxStyle style = item.Style;
        double availableWidth = Math.Max(1D, containingWidth - style.MarginLeft - style.MarginRight);
        double boxWidth = ResolveBoxWidth(availableWidth, style);
        double contentWidth = Math.Max(1D, boxWidth - style.HorizontalInsets);
        var run = new HtmlInlineRun(ApplyTextTransform(item.AnonymousText, style.TextTransform), style, item.Link, item.Source);
        HtmlInlineLayout inline = LayoutInlineRuns(new[] { run }, contentWidth, style);
        double boxHeight = ResolveBoxHeight(inline.Height, style);
        double outerHeight = Math.Max(0.01D, style.MarginTop + boxHeight + style.MarginBottom);
        var visuals = new List<HtmlRenderVisual>();
        if (item.PaintAnonymousBox) AddGeneratedBoxPaint(visuals, style, style.MarginLeft, style.MarginTop, boxWidth, boxHeight, item.SourceElement, item.Source);
        double contentX = style.MarginLeft + style.BorderWidth + style.PaddingLeft;
        double contentY = style.MarginTop + style.BorderWidth + style.PaddingTop;
        foreach (HtmlRenderVisual visual in inline.Visuals) visuals.Add(visual.Translate(contentX, contentY, visuals.Count));
        if (item.PaintAnonymousBox) AddGeneratedBoxOutlinePaint(visuals, style, style.MarginLeft, style.MarginTop, boxWidth, boxHeight, item.SourceElement, item.Source);
        IEnumerable<double> breakOffsets = inline.BreakOffsets.Select(offset => contentY + offset).Concat(new[] { outerHeight });
        var block = new HtmlRenderFlowBlock(
            containingWidth,
            outerHeight,
            visuals,
            style.BreakBefore,
            style.BreakAfter,
            style.AvoidBreakInside,
            item.Source,
            breakOffsets,
            inline.BreakOffsets.Select(offset => contentY + offset),
            style.Orphans,
            style.Widows,
            pageName: style.PageName ?? parentStyle.PageName);
        return ApplyPositioning(block, style, containingWidth, ResolveContainingBlockHeight(parentStyle), item.Source);
    }

    private string? ResolveFlexItemLink(IElement element) {
        for (IElement? current = element; current != null; current = current.ParentElement) {
            if (string.Equals(current.TagName, "a", StringComparison.OrdinalIgnoreCase)) {
                return ResolveSafeLink(current.GetAttribute("href"), current);
            }
        }

        return null;
    }

    private static HtmlRenderBoxStyle CreateAnonymousFlexStyle(HtmlRenderBoxStyle parentStyle) => new HtmlRenderBoxStyle {
        Display = "block",
        Font = parentStyle.Font,
        Color = parentStyle.Color,
        Alignment = parentStyle.Alignment,
        LineHeight = parentStyle.LineHeight,
        PreserveWhitespace = parentStyle.PreserveWhitespace,
        TextTransform = parentStyle.TextTransform,
        SemanticRole = "anonymous-flex-item",
        Orphans = parentStyle.Orphans,
        Widows = parentStyle.Widows,
        PageName = parentStyle.PageName
    };

    private static HtmlRenderBoxStyle BlockifyFlexItemStyle(HtmlRenderBoxStyle style) {
        string display;
        switch (style.Display) {
            case "inline-flex": display = "flex"; break;
            case "inline-grid": display = "grid"; break;
            case "inline":
            case "inline-block": display = "block"; break;
            default: return style;
        }

        HtmlRenderBoxStyle blockified = style.Clone();
        blockified.Display = display;
        return blockified;
    }
}
