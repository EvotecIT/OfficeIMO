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
        bool wrapping = style.FlexWrap != "nowrap";
        if (style.UnsupportedRowGap.Length > 0) ReportUnsupportedFlexValue(element, "row-gap=" + style.UnsupportedRowGap);
        if (wrapping && style.UnsupportedColumnGap.Length > 0) ReportUnsupportedFlexValue(element, "column-gap=" + style.UnsupportedColumnGap);

        List<FlexItem> orderedItems = items.OrderBy(item => item.Style.Order).ThenBy(item => item.SourceIndex).ToList();
        foreach (FlexItem item in orderedItems) {
            CheckCancellation();
            item.HasExplicitCrossSize = item.Style.ExplicitWidth.HasValue;
            item.CrossBasis = ResolveColumnFlexCrossBasis(item, contentWidth);
            string alignment = ResolveFlexAlignment(item.Style.AlignSelf, style.AlignItems);
            double initialCrossSize = !wrapping && alignment == "stretch" && !HasColumnFlexCrossAutoMargin(item.Style) ? contentWidth : item.CrossBasis;
            ApplyColumnFlexCrossSize(item, initialCrossSize);
            item.InitialCrossSize = initialCrossSize;
            item.Block = LayoutFlexItem(item, Math.Max(1D, initialCrossSize), style, depth + 1);
        }

        bool hasDefiniteHeight = style.ExplicitHeight.HasValue;
        double heightReference = hasDefiniteHeight
            ? Math.Max(0D, style.BorderBox ? style.ExplicitHeight!.Value - style.VerticalInsets : style.ExplicitHeight!.Value)
            : 0D;
        foreach (FlexItem item in orderedItems) item.Basis = ResolveColumnFlexBasis(item, heightReference, hasDefiniteHeight);

        double mainGap = orderedItems.Count > 1 ? style.RowGap : 0D;
        double naturalContentHeight = orderedItems.Sum(item => item.Basis) + mainGap * Math.Max(0, orderedItems.Count - 1);
        double boxHeight = ResolveBoxHeight(naturalContentHeight, style);
        double contentHeight = Math.Max(0D, boxHeight - style.VerticalInsets);
        string effectiveWrap = wrapping && hasDefiniteHeight ? style.FlexWrap : "nowrap";
        List<FlexLine> lines = CreateFlexLines(orderedItems, effectiveWrap, contentHeight, mainGap);
        foreach (FlexLine line in lines) {
            CheckCancellation();
            double availableForItems = Math.Max(0D, contentHeight - mainGap * Math.Max(0, line.Items.Count - 1));
            ResolveFlexMainSizes(line.Items, availableForItems, vertical: true);
            line.CrossSize = line.Items.Count == 0 ? 0D : line.Items.Max(item => item.CrossBasis);
        }

        string source = HtmlRenderStyleResolver.DescribeSource(element);
        double crossGap = lines.Count > 1 ? style.ColumnGap : 0D;
        ResolveFlexLineOffsets(lines, style, contentWidth, crossGap, source);
        foreach (FlexLine line in lines) {
            CheckCancellation();
            ResolveFlexMainOffsets(line.Items, style, contentHeight, mainGap, style.FlexDirection == "column-reverse", vertical: true, source: source);
            foreach (FlexItem item in line.Items) {
                string alignment = ResolveFlexAlignment(item.Style.AlignSelf, style.AlignItems);
                double targetCrossSize = alignment == "stretch" && !item.HasExplicitCrossSize && !HasColumnFlexCrossAutoMargin(item.Style) ? line.CrossSize : item.CrossBasis;
                bool mainSizeWasDefinite = item.Style.ExplicitHeight.HasValue;
                bool canReuseInitialLayout = Math.Abs(targetCrossSize - item.InitialCrossSize) <= 0.0001D
                    && Math.Abs(line.CrossSize - item.InitialCrossSize) <= 0.0001D
                    && Math.Abs(item.MainSize - item.Block!.Height) <= 0.0001D
                    && (!hasDefiniteHeight || mainSizeWasDefinite);
                ApplyColumnFlexCrossSize(item, targetCrossSize);
                ApplyColumnFlexMainSize(item);
                if (!canReuseInitialLayout) item.Block = LayoutFlexItem(item, Math.Max(1D, line.CrossSize), style, depth + 1);
                item.CrossOffset = ResolveColumnFlexCrossOffset(item, style, line.CrossSize);
            }
        }

        double outerHeight = Math.Max(0.01D, style.MarginTop + boxHeight + style.MarginBottom);
        var visuals = new List<HtmlRenderVisual>();
        AddBoxPaint(visuals, style, style.MarginLeft, style.MarginTop, boxWidth, boxHeight, element);
        AppendLocalPositionedVisuals(
            element,
            Math.Max(1D, boxWidth - style.BorderLeftWidth - style.BorderRightWidth),
            Math.Max(0.01D, boxHeight - style.BorderTopWidth - style.BorderBottomWidth),
            style.MarginLeft + style.BorderLeftWidth,
            style.MarginTop + style.BorderTopWidth,
            PositionedPaintBand.Negative,
            visuals);
        double contentX = style.MarginLeft + style.BorderLeftWidth + style.PaddingLeft;
        double contentY = style.MarginTop + style.BorderTopWidth + style.PaddingTop;
        var itemPaintLayers = new List<FlowPaintLayer>();
        foreach (FlexLine line in lines) {
            CheckCancellation();
            foreach (FlexItem item in line.Items) {
                if (item.Element != null) {
                    RecordNormalFlowPlacement(item.Element, element, line.CrossOffset + item.CrossOffset, item.MainOffset, item.Style);
                }
                itemPaintLayers.Add(new FlowPaintLayer(
                    item.Block!,
                    contentX + line.CrossOffset + item.CrossOffset,
                    contentY + item.MainOffset,
                    itemPaintLayers.Count));
            }
        }
        AppendFlowPaintLayers(visuals, itemPaintLayers);
        AppendLocalPositionedVisuals(
            element,
            Math.Max(1D, boxWidth - style.BorderLeftWidth - style.BorderRightWidth),
            Math.Max(0.01D, boxHeight - style.BorderTopWidth - style.BorderBottomWidth),
            style.MarginLeft + style.BorderLeftWidth,
            style.MarginTop + style.BorderTopWidth,
            PositionedPaintBand.NonNegative,
            visuals);
        AddBoxOutlinePaint(visuals, style, style.MarginLeft, style.MarginTop, boxWidth, boxHeight, element);

        IEnumerable<double>? breakOffsets = lines.Count != 1
            ? null
            : lines[0].Items.SelectMany(item =>
                    new[] { contentY + item.MainOffset }
                        .Concat(item.Block!.BreakOffsets.Select(offset => contentY + item.MainOffset + offset)))
                .Distinct()
                .OrderBy(offset => offset);
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

    private double ResolveColumnFlexCrossBasis(FlexItem item, double contentWidth) {
        HtmlRenderBoxStyle style = item.Style;
        string tag = item.TagName;
        if (tag == "img" && item.Element != null) {
            double imageOuter = ResolveReplacedImageBoxWidth(item.Element, style) + style.MarginLeft + style.MarginRight;
            return Math.Max(1D, Math.Min(contentWidth, imageOuter));
        }
        if (style.ExplicitWidth.HasValue) return ResolveColumnFlexOuterWidth(style, contentWidth);
        if (tag == "table") return contentWidth;
        double boxBasis;
        string content = CollapseFlexText(item.TextContent);
        double measured = content.Length == 0 ? 1D : MeasureText(ApplyTextTransform(content, style.TextTransform), style.Font);
        boxBasis = measured + style.HorizontalInsets;

        if (style.MinWidth.HasValue) boxBasis = Math.Max(boxBasis, style.MinWidth.Value + (style.BorderBox ? 0D : style.HorizontalInsets));
        if (style.MaxWidth.HasValue) boxBasis = Math.Min(boxBasis, style.MaxWidth.Value + (style.BorderBox ? 0D : style.HorizontalInsets));
        double outer = boxBasis + style.MarginLeft + style.MarginRight;
        return Math.Max(1D, Math.Min(contentWidth, outer));
    }

    private double ResolveColumnFlexBasis(FlexItem item, double heightReference, bool hasDefiniteHeight) {
        string basis = item.Style.FlexBasis;
        if (basis == "auto" || !hasDefiniteHeight && basis.IndexOf('%') >= 0) return item.Block!.Height;
        if (HtmlRenderCssValues.TryLength(basis, heightReference, item.Style.Font.Size, _options.DefaultFontSize, out double resolved)) {
            double boxBasis = Math.Max(0D, resolved) + (item.Style.BorderBox ? 0D : item.Style.VerticalInsets);
            return boxBasis + item.Style.MarginTop + item.Style.MarginBottom;
        }

        ReportUnsupportedFlexValue(item, "flex-basis=" + basis);
        return item.Block!.Height;
    }

    private static void ApplyColumnFlexCrossSize(FlexItem item, double targetOuterWidth) {
        if (item.HasExplicitCrossSize) return;
        HtmlRenderBoxStyle style = item.Style.Clone();
        double targetBoxWidth = Math.Max(0.01D, targetOuterWidth - style.MarginLeft - style.MarginRight);
        style.ExplicitWidth = style.BorderBox
            ? targetBoxWidth
            : Math.Max(0.01D, targetBoxWidth - style.HorizontalInsets);
        item.Style = style;
    }

    private static void ApplyColumnFlexMainSize(FlexItem item) {
        HtmlRenderBoxStyle style = item.Style.Clone();
        double targetBoxHeight = Math.Max(0.01D, item.MainSize - style.MarginTop - style.MarginBottom);
        style.ExplicitHeight = style.BorderBox
            ? targetBoxHeight
            : Math.Max(0.01D, targetBoxHeight - style.VerticalInsets);
        item.Style = style;
    }

    private double ResolveColumnFlexCrossOffset(FlexItem item, HtmlRenderBoxStyle containerStyle, double lineCrossSize) {
        string alignment = ResolveFlexAlignment(item.Style.AlignSelf, containerStyle.AlignItems);
        double outerWidth = ResolveColumnFlexOuterWidth(item.Style, lineCrossSize);
        double remaining = Math.Max(0D, lineCrossSize - outerWidth);
        if (item.Style.MarginLeftAuto || item.Style.MarginRightAuto) {
            if (item.Style.MarginLeftAuto && item.Style.MarginRightAuto) return remaining / 2D;
            return item.Style.MarginLeftAuto ? remaining : 0D;
        }
        bool reverse = containerStyle.FlexWrap == "wrap-reverse";
        if (alignment == "flex-end") return reverse ? 0D : remaining;
        if (alignment == "end") return remaining;
        if (alignment == "center") return remaining / 2D;
        if (alignment == "stretch" || alignment == "flex-start") return reverse ? remaining : 0D;
        if (alignment == "start") return 0D;
        ReportUnsupportedFlexValue(item, "align-self=" + alignment);
        return 0D;
    }

    private double ResolveColumnFlexOuterWidth(HtmlRenderBoxStyle style, double availableCrossSize) {
        double available = Math.Max(1D, availableCrossSize - style.MarginLeft - style.MarginRight);
        return style.MarginLeft + ResolveBoxWidth(available, style) + style.MarginRight;
    }

    private static bool HasColumnFlexCrossAutoMargin(HtmlRenderBoxStyle style) => style.MarginLeftAuto || style.MarginRightAuto;
}
