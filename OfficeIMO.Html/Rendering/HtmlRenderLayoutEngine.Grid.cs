using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private bool TryLayoutGridContainer(
        IElement element,
        double containingWidth,
        HtmlRenderBoxStyle style,
        int depth,
        out HtmlRenderFlowBlock block) {
        block = null!;
        if (!TryCollectFlexItems(element, containingWidth, style, depth, out List<FlexItem> formattingItems)) return false;
        string source = HtmlRenderStyleResolver.DescribeSource(element);
        double availableWidth = Math.Max(1D, containingWidth - style.MarginLeft - style.MarginRight);
        double boxWidth = ResolveBoxWidth(availableWidth, style);
        double contentWidth = Math.Max(1D, boxWidth - style.HorizontalInsets);
        if (style.UnsupportedColumnGap.Length > 0) ReportUnsupportedGridValue(source, "column-gap=" + style.UnsupportedColumnGap);
        if (style.UnsupportedRowGap.Length > 0) ReportUnsupportedGridValue(source, "row-gap=" + style.UnsupportedRowGap);

        double? declaredContentHeight = ResolveGridDeclaredContentHeight(style);
        List<GridTrack> columnTracks = ParseGridTracks(style.GridTemplateColumns, contentWidth, percentageReferenceIsDefinite: true, style, source, "grid-template-columns");
        List<GridTrack> rowTracks = ParseGridTracks(
            style.GridTemplateRows,
            declaredContentHeight ?? 0D,
            declaredContentHeight.HasValue,
            style,
            source,
            "grid-template-rows");
        IReadOnlyDictionary<string, GridAreaDefinition> areas = ParseGridTemplateAreas(style.GridTemplateAreas, source, out int areaRowCount, out int areaColumnCount);
        IReadOnlyDictionary<string, int> columnLineNames = ParseGridLineNames(style.GridTemplateColumns);
        IReadOnlyDictionary<string, int> rowLineNames = ParseGridLineNames(style.GridTemplateRows);
        int explicitColumnCount = Math.Max(1, Math.Max(columnTracks.Count, areaColumnCount));
        int explicitRowCount = Math.Max(1, Math.Max(rowTracks.Count, areaRowCount));
        List<GridItem> items = PlaceGridItems(formattingItems, explicitColumnCount, explicitRowCount, style, source, areas, columnLineNames, rowLineNames, out int columnCount, out int rowCount);
        CollapseTrailingAutoFitColumns(style, items, columnTracks, ref columnCount);
        rowCount = Math.Max(rowCount, Math.Max(1, areaRowCount));
        EnsureGridTrackCount(columnTracks, columnCount, style.GridAutoColumns, contentWidth, percentageReferenceIsDefinite: true, style, source, "grid-auto-columns");
        double columnGap = columnCount > 1 ? style.ColumnGap : 0D;
        List<double> columnSizes = ResolveGridTrackSizes(columnTracks, contentWidth, columnGap);
        GridAxisLayout columns = ResolveGridAxisLayout(columnTracks, columnSizes, contentWidth, columnGap, style.JustifyContent, source, "justify-content");

        foreach (GridItem item in items) {
            CheckCancellation();
            double cellWidth = columns.SpanSize(item.Column, item.ColumnSpan);
            ApplyInitialGridItemWidth(item, style, cellWidth);
            item.Block = LayoutFlexItem(item.Item, Math.Max(1D, cellWidth), style, depth + 1);
        }

        EnsureGridTrackCount(
            rowTracks,
            rowCount,
            style.GridAutoRows,
            declaredContentHeight ?? 0D,
            declaredContentHeight.HasValue,
            style,
            source,
            "grid-auto-rows");
        double rowGap = rowCount > 1 ? style.RowGap : 0D;
        List<double> rowSizes = ResolveNaturalGridRows(rowTracks, items, rowGap, declaredContentHeight);
        double naturalContentHeight = rowSizes.Sum() + rowGap * Math.Max(0, rowCount - 1);
        double boxHeight = ResolveBoxHeight(naturalContentHeight, style);
        double contentHeight = Math.Max(0D, boxHeight - style.VerticalInsets);
        GridAxisLayout rows = ResolveGridAxisLayout(rowTracks, rowSizes, contentHeight, rowGap, style.AlignContent, source, "align-content");
        RecordGridPositionedContainingRects(
            element,
            style,
            contentWidth,
            contentHeight,
            columns,
            rows,
            areas,
            columnLineNames,
            rowLineNames);

        foreach (GridItem item in items) {
            CheckCancellation();
            double cellWidth = columns.SpanSize(item.Column, item.ColumnSpan);
            double cellHeight = rows.SpanSize(item.Row, item.RowSpan);
            ApplyFinalGridItemSize(item, style, cellWidth, cellHeight);
            item.Block = LayoutFlexItem(item.Item, Math.Max(1D, cellWidth), style, depth + 1);
            item.OffsetX = ResolveGridHorizontalOffset(item, style, cellWidth);
            item.OffsetY = ResolveGridVerticalOffset(item, style, cellHeight);
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
        foreach (GridItem item in items) {
            CheckCancellation();
            double itemX = contentX + columns.Positions[item.Column] + item.OffsetX;
            double itemY = contentY + rows.Positions[item.Row] + item.OffsetY;
            if (item.Item.Element != null) {
                RecordNormalFlowPlacement(
                    item.Item.Element,
                    element,
                    columns.Positions[item.Column] + item.OffsetX,
                    rows.Positions[item.Row] + item.OffsetY,
                    item.Item.Style);
            }
            itemPaintLayers.Add(new FlowPaintLayer(item.Block!, itemX, itemY, itemPaintLayers.Count));
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

        IEnumerable<double> rowBreakOffsets = Enumerable.Range(1, Math.Max(0, rowCount - 1))
            .Where(boundary => !items.Any(item => item.Row < boundary && item.Row + item.RowSpan > boundary))
            .Select(boundary => contentY + rows.Positions[boundary]);
        var rowItemCountDeltas = new int[rowCount + 1];
        foreach (GridItem item in items) {
            rowItemCountDeltas[item.Row]++;
            rowItemCountDeltas[Math.Min(rowCount, item.Row + item.RowSpan)]--;
        }
        var rowItemCounts = new int[rowCount];
        int activeRowItems = 0;
        for (int row = 0; row < rowCount; row++) {
            activeRowItems += rowItemCountDeltas[row];
            rowItemCounts[row] = activeRowItems;
        }
        IEnumerable<double> itemBreakOffsets = items
            .Where(item => item.RowSpan == 1 && rowItemCounts[item.Row] == 1)
            .SelectMany(item => item.Block!.BreakOffsets.Select(offset =>
                contentY + rows.Positions[item.Row] + item.OffsetY + offset));
        IEnumerable<double> breakOffsets = rowBreakOffsets.Concat(itemBreakOffsets).Distinct().OrderBy(offset => offset);
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

    private double? ResolveGridDeclaredContentHeight(HtmlRenderBoxStyle style) {
        if (!style.ExplicitHeight.HasValue) return null;
        return style.BorderBox
            ? Math.Max(0D, style.ExplicitHeight.Value - style.VerticalInsets)
            : Math.Max(0D, style.ExplicitHeight.Value);
    }

    private static List<double> ResolveNaturalGridRows(
        IReadOnlyList<GridTrack> tracks,
        IReadOnlyList<GridItem> items,
        double gap,
        double? declaredContentHeight) {
        var sizes = tracks.Select(track => track.Kind == GridTrackKind.Fixed ? Math.Max(track.Value, track.Minimum) : track.Minimum).ToList();
        foreach (GridItem item in items.OrderBy(item => item.RowSpan)) {
            double required = item.Block!.Height;
            double current = sizes.Skip(item.Row).Take(item.RowSpan).Sum() + gap * Math.Max(0, item.RowSpan - 1);
            double deficit = Math.Max(0D, required - current);
            if (deficit <= 0D) continue;
            List<int> flexible = Enumerable.Range(item.Row, item.RowSpan).Where(index => tracks[index].Kind != GridTrackKind.Fixed).ToList();
            if (flexible.Count == 0) flexible.AddRange(Enumerable.Range(item.Row, item.RowSpan));
            double addition = deficit / flexible.Count;
            foreach (int index in flexible) sizes[index] += addition;
        }

        if (declaredContentHeight.HasValue) {
            double trackSpace = Math.Max(0D, declaredContentHeight.Value - gap * Math.Max(0, tracks.Count - 1));
            double fractions = tracks.Where(track => track.Kind == GridTrackKind.Fraction).Sum(track => track.Value);
            if (fractions > 0D) DistributeGridFractions(tracks, sizes, trackSpace);
        }
        return sizes;
    }

    private void ApplyInitialGridItemWidth(GridItem item, HtmlRenderBoxStyle containerStyle, double cellWidth) {
        string alignment = ResolveGridAlignment(item.Item.Style.JustifySelf, containerStyle.JustifyItems);
        double targetWidth = alignment == "stretch" && !item.HasExplicitWidth && !HasHorizontalAutoMargin(item.Item.Style)
            ? cellWidth
            : ResolveColumnFlexCrossBasis(item.Item, cellWidth);
        ApplyGridItemWidth(item, targetWidth);
    }

    private void ApplyFinalGridItemSize(GridItem item, HtmlRenderBoxStyle containerStyle, double cellWidth, double cellHeight) {
        HtmlRenderBoxStyle style = item.Item.Style.Clone();
        string horizontal = ResolveGridAlignment(style.JustifySelf, containerStyle.JustifyItems);
        string vertical = ResolveGridAlignment(style.AlignSelf, containerStyle.AlignItems);
        if (horizontal == "stretch" && !item.HasExplicitWidth && !HasHorizontalAutoMargin(style)) {
            double targetBoxWidth = Math.Max(0.01D, cellWidth - style.MarginLeft - style.MarginRight);
            style.ExplicitWidth = style.BorderBox ? targetBoxWidth : Math.Max(0.01D, targetBoxWidth - style.HorizontalInsets);
        }
        if (vertical == "stretch" && !item.HasExplicitHeight && !HasVerticalAutoMargin(style)) {
            double targetBoxHeight = Math.Max(0.01D, cellHeight - style.MarginTop - style.MarginBottom);
            style.ExplicitHeight = style.BorderBox ? targetBoxHeight : Math.Max(0.01D, targetBoxHeight - style.VerticalInsets);
        }
        item.Item.Style = style;
    }

    private static void ApplyGridItemWidth(GridItem item, double targetOuterWidth) {
        if (item.HasExplicitWidth) return;
        HtmlRenderBoxStyle style = item.Item.Style.Clone();
        double targetBoxWidth = Math.Max(0.01D, targetOuterWidth - style.MarginLeft - style.MarginRight);
        style.ExplicitWidth = style.BorderBox ? targetBoxWidth : Math.Max(0.01D, targetBoxWidth - style.HorizontalInsets);
        item.Item.Style = style;
    }

    private double ResolveGridHorizontalOffset(GridItem item, HtmlRenderBoxStyle containerStyle, double cellWidth) {
        HtmlRenderBoxStyle style = item.Item.Style;
        double outerWidth = ResolveColumnFlexOuterWidth(style, cellWidth);
        double remaining = Math.Max(0D, cellWidth - outerWidth);
        if (style.MarginLeftAuto || style.MarginRightAuto) {
            if (style.MarginLeftAuto && style.MarginRightAuto) return remaining / 2D;
            return style.MarginLeftAuto ? remaining : 0D;
        }
        return ResolveGridAlignmentOffset(ResolveGridAlignment(style.JustifySelf, containerStyle.JustifyItems), remaining, item.Item.Source, "justify-self");
    }

    private double ResolveGridVerticalOffset(GridItem item, HtmlRenderBoxStyle containerStyle, double cellHeight) {
        HtmlRenderBoxStyle style = item.Item.Style;
        double remaining = Math.Max(0D, cellHeight - item.Block!.Height);
        if (style.MarginTopAuto || style.MarginBottomAuto) {
            if (style.MarginTopAuto && style.MarginBottomAuto) return remaining / 2D;
            return style.MarginTopAuto ? remaining : 0D;
        }
        return ResolveGridAlignmentOffset(ResolveGridAlignment(style.AlignSelf, containerStyle.AlignItems), remaining, item.Item.Source, "align-self");
    }

    private double ResolveGridAlignmentOffset(string alignment, double remaining, string source, string property) {
        if (alignment == "end" || alignment == "flex-end") return remaining;
        if (alignment == "center") return remaining / 2D;
        if (alignment == "stretch" || alignment == "start" || alignment == "flex-start") return 0D;
        ReportUnsupportedGridValue(source, property + "=" + alignment);
        return 0D;
    }

    private static string ResolveGridAlignment(string self, string container) {
        string resolved = self == "auto" ? container : self;
        return resolved == "normal" ? "stretch" : resolved;
    }

    private static bool HasHorizontalAutoMargin(HtmlRenderBoxStyle style) => style.MarginLeftAuto || style.MarginRightAuto;
    private static bool HasVerticalAutoMargin(HtmlRenderBoxStyle style) => style.MarginTopAuto || style.MarginBottomAuto;
}
