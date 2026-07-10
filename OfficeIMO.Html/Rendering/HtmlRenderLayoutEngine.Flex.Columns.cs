using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private bool TryLayoutColumnFlexContainer(
        IElement element,
        double containingWidth,
        HtmlRenderBoxStyle style,
        int depth,
        List<FlexItem> items,
        out HtmlRenderFlowBlock block) {
        double availableWidth = Math.Max(1D, containingWidth - style.MarginLeft - style.MarginRight);
        double boxWidth = ResolveBoxWidth(availableWidth, style);
        double contentWidth = Math.Max(1D, boxWidth - style.HorizontalInsets);
        if (style.UnsupportedRowGap.Length > 0) ReportUnsupportedFlexValue(element, "row-gap=" + style.UnsupportedRowGap);

        List<FlexItem> orderedItems = items.OrderBy(item => item.Style.Order).ThenBy(item => item.SourceIndex).ToList();
        foreach (FlexItem item in orderedItems) {
            StretchColumnFlexItem(item, style, contentWidth);
            item.Block = LayoutElement(item.Element, contentWidth, item.Style, style, depth + 1);
        }

        bool hasDefiniteHeight = style.ExplicitHeight.HasValue;
        double heightReference = hasDefiniteHeight
            ? Math.Max(0D, style.BorderBox ? style.ExplicitHeight!.Value - style.VerticalInsets : style.ExplicitHeight!.Value)
            : 0D;
        foreach (FlexItem item in orderedItems) item.Basis = ResolveColumnFlexBasis(item, heightReference, hasDefiniteHeight);

        double gap = orderedItems.Count > 1 ? style.RowGap : 0D;
        double naturalContentHeight = orderedItems.Sum(item => item.Basis) + gap * Math.Max(0, orderedItems.Count - 1);
        double boxHeight = ResolveBoxHeight(naturalContentHeight, style);
        double contentHeight = Math.Max(0D, boxHeight - style.VerticalInsets);
        double availableForItems = Math.Max(0D, contentHeight - gap * Math.Max(0, orderedItems.Count - 1));
        ResolveFlexMainSizes(orderedItems, availableForItems, vertical: true);
        foreach (FlexItem item in orderedItems) {
            ApplyColumnFlexMainSize(item);
            item.Block = LayoutElement(item.Element, contentWidth, item.Style, style, depth + 1);
        }

        string source = HtmlRenderStyleResolver.DescribeSource(element);
        ResolveFlexMainOffsets(orderedItems, style, contentHeight, gap, style.FlexDirection == "column-reverse", source);
        foreach (FlexItem item in orderedItems) item.CrossOffset = ResolveColumnFlexCrossOffset(item, style, contentWidth);

        double outerHeight = Math.Max(0.01D, style.MarginTop + boxHeight + style.MarginBottom);
        var visuals = new List<HtmlRenderVisual>();
        AddBoxPaint(visuals, style, style.MarginLeft, style.MarginTop, boxWidth, boxHeight, element);
        double contentX = style.MarginLeft + style.BorderWidth + style.PaddingLeft;
        double contentY = style.MarginTop + style.BorderWidth + style.PaddingTop;
        foreach (FlexItem item in orderedItems) {
            foreach (HtmlRenderVisual visual in item.Block!.Visuals) {
                visuals.Add(visual.Translate(contentX + item.CrossOffset, contentY + item.MainOffset, visuals.Count));
            }
        }

        IEnumerable<double>? breakOffsets = orderedItems.Count < 2
            ? null
            : orderedItems.Select(item => contentY + item.MainOffset)
                .Distinct()
                .OrderBy(offset => offset)
                .Skip(1);
        block = new HtmlRenderFlowBlock(
            containingWidth,
            outerHeight,
            visuals,
            style.BreakBefore,
            style.BreakAfter,
            style.AvoidBreakInside,
            source,
            breakOffsets,
            pageName: style.PageName);
        return true;
    }

    private void StretchColumnFlexItem(FlexItem item, HtmlRenderBoxStyle containerStyle, double contentWidth) {
        string alignment = ResolveFlexAlignment(item.Style.AlignSelf, containerStyle.AlignItems);
        if (alignment != "stretch" || item.Style.ExplicitWidth.HasValue) return;
        double targetBoxWidth = Math.Max(0.01D, contentWidth - item.Style.MarginLeft - item.Style.MarginRight);
        var stretched = item.Style.Clone();
        stretched.ExplicitWidth = stretched.BorderBox
            ? targetBoxWidth
            : Math.Max(0.01D, targetBoxWidth - stretched.HorizontalInsets);
        item.Style = stretched;
    }

    private double ResolveColumnFlexBasis(FlexItem item, double heightReference, bool hasDefiniteHeight) {
        string basis = item.Style.FlexBasis;
        if (basis == "auto" || !hasDefiniteHeight && basis.IndexOf('%') >= 0) return item.Block!.Height;
        if (HtmlRenderCssValues.TryLength(basis, heightReference, item.Style.Font.Size, _options.DefaultFontSize, out double resolved)) {
            double boxBasis = Math.Max(0D, resolved) + (item.Style.BorderBox ? 0D : item.Style.VerticalInsets);
            return boxBasis + item.Style.MarginTop + item.Style.MarginBottom;
        }

        ReportUnsupportedFlexValue(item.Element, "flex-basis=" + basis);
        return item.Block!.Height;
    }

    private static void ApplyColumnFlexMainSize(FlexItem item) {
        HtmlRenderBoxStyle style = item.Style.Clone();
        double targetBoxHeight = Math.Max(0.01D, item.MainSize - style.MarginTop - style.MarginBottom);
        style.ExplicitHeight = style.BorderBox
            ? targetBoxHeight
            : Math.Max(0.01D, targetBoxHeight - style.VerticalInsets);
        item.Style = style;
    }

    private double ResolveColumnFlexCrossOffset(FlexItem item, HtmlRenderBoxStyle containerStyle, double contentWidth) {
        string alignment = ResolveFlexAlignment(item.Style.AlignSelf, containerStyle.AlignItems);
        double outerWidth = ResolveColumnFlexOuterWidth(item.Style, contentWidth);
        double remaining = Math.Max(0D, contentWidth - outerWidth);
        if (alignment == "flex-end" || alignment == "end") return remaining;
        if (alignment == "center") return remaining / 2D;
        if (alignment == "stretch" || alignment == "flex-start" || alignment == "start") return 0D;
        ReportUnsupportedFlexValue(item.Element, "align-self=" + alignment);
        return 0D;
    }

    private double ResolveColumnFlexOuterWidth(HtmlRenderBoxStyle style, double contentWidth) {
        double available = Math.Max(1D, contentWidth - style.MarginLeft - style.MarginRight);
        return style.MarginLeft + ResolveBoxWidth(available, style) + style.MarginRight;
    }
}
