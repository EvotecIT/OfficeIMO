using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private void AddInlineGridRun(
        IElement element,
        double availableWidth,
        HtmlRenderBoxStyle parentStyle,
        int depth,
        HtmlRenderBoxStyle inlineStyle,
        string? link,
        double inheritedPaintOffsetX,
        double inheritedPaintOffsetY,
        ICollection<HtmlInlineRun> runs) {
        HtmlRenderBoxStyle gridStyle = BlockifyFlexItemStyle(inlineStyle);
        double outerWidth = ResolveInlineGridWidth(element, availableWidth, gridStyle, depth + 1);
        if (!gridStyle.ExplicitWidth.HasValue) {
            gridStyle = gridStyle.Clone();
            double targetBoxWidth = Math.Max(0.01D, outerWidth - gridStyle.MarginLeft - gridStyle.MarginRight);
            gridStyle.ExplicitWidth = gridStyle.BorderBox
                ? targetBoxWidth
                : Math.Max(0.01D, targetBoxWidth - gridStyle.HorizontalInsets);
        }

        HtmlRenderFlowBlock atomic = LayoutElement(element, outerWidth, gridStyle, parentStyle, depth + 1);
        runs.Add(new HtmlInlineRun(
            atomic,
            inlineStyle,
            link,
            HtmlRenderStyleResolver.DescribeSource(element),
            inheritedPaintOffsetX,
            inheritedPaintOffsetY,
            element));
    }

    private double ResolveInlineGridWidth(IElement element, double availableWidth, HtmlRenderBoxStyle style, int depth) =>
        ResolveIntrinsicGridWidths(element, availableWidth, style, depth).Maximum;

    private (double Minimum, double Maximum) ResolveIntrinsicGridWidths(IElement element, double availableWidth, HtmlRenderBoxStyle style, int depth, bool clampToAvailable = true) {
        double availableOuterWidth = Math.Max(1D, availableWidth);
        double availableBoxWidth = Math.Max(1D, availableOuterWidth - style.MarginLeft - style.MarginRight);
        if (style.ExplicitWidth.HasValue) {
            double definite = clampToAvailable
                ? Math.Min(availableOuterWidth, style.MarginLeft + ResolveBoxWidth(availableBoxWidth, style) + style.MarginRight)
                : ResolveGridMeasuredContribution(style, style.ExplicitWidth.Value - (style.BorderBox ? style.HorizontalInsets : 0D));
            return (definite, definite);
        }

        if (!TryCollectFlexItems(element, availableOuterWidth, style, depth, captureRunningElements: false,
            out List<FlexItem> formattingItems, out _, registerOutOfFlowElements: false)) return (availableOuterWidth, availableOuterWidth);
        string source = HtmlRenderStyleResolver.DescribeSource(element);
        List<GridTrack> tracks = ParseGridTracks(style.GridTemplateColumns, availableBoxWidth, percentageReferenceIsDefinite: true, style, source, "grid-template-columns");
        double? declaredContentHeight = ResolveGridDeclaredContentHeight(style);
        List<GridTrack> rows = ParseGridTracks(style.GridTemplateRows, declaredContentHeight ?? 0D, declaredContentHeight.HasValue, style, source, "grid-template-rows");
        IReadOnlyDictionary<string, GridAreaDefinition> areas = ParseGridTemplateAreas(style.GridTemplateAreas, source, out int areaRowCount, out int areaColumnCount);
        IReadOnlyDictionary<string, int> columnLineNames = ParseGridLineNames(style.GridTemplateColumns);
        IReadOnlyDictionary<string, int> rowLineNames = ParseGridLineNames(style.GridTemplateRows);
        int explicitColumns = Math.Max(tracks.Count, areaColumnCount);
        int explicitRows = Math.Max(rows.Count, areaRowCount);
        columnLineNames = AddGridAreaLineNames(columnLineNames, areas, rows: false);
        rowLineNames = AddGridAreaLineNames(rowLineNames, areas, rows: true);
        List<GridItem> items = PlaceGridItems(formattingItems, explicitColumns, explicitRows, style, source, columnLineNames, rowLineNames,
            out int columnCount, out _, out int leadingColumnCount, out int leadingRowCount);
        PrependImplicitGridTracks(tracks, leadingColumnCount, style.GridAutoColumns, availableBoxWidth, true, style, source, "grid-auto-columns");
        columnLineNames = OffsetGridLineNames(columnLineNames, leadingColumnCount);
        rowLineNames = OffsetGridLineNames(rowLineNames, leadingRowCount);
        CollapseEmptyAutoFitColumns(style, items, tracks, ref columnCount);
        EnsureGridTrackCount(tracks, columnCount, style.GridAutoColumns, availableBoxWidth, percentageReferenceIsDefinite: true, style, source, "grid-auto-columns");
        List<GridIntrinsicContribution> contributions = CollectGridIntrinsicContributions(items, availableBoxWidth, style.ColumnGap, columnLineNames, rowLineNames, depth);
        var measured = new Dictionary<FlexItem, (double Minimum, double Maximum)>();
        List<double> maximumSizes = ResolveGridIntrinsicTrackBases(tracks, contributions, availableBoxWidth, style.ColumnGap,
            includeFractionTracks: true, depth: depth, measurements: measured);
        List<double> minimumSizes = ResolveGridIntrinsicTrackBases(tracks, contributions, availableBoxWidth, style.ColumnGap,
            includeFractionTracks: true, depth: depth, minimumContribution: true, measurements: measured);
        return (ResolveWidth(minimumSizes), ResolveWidth(maximumSizes));

        double ResolveWidth(IReadOnlyList<double> sizes) {
            double intrinsicContentWidth = sizes.Sum() + style.ColumnGap * CountGridBaseGaps(tracks);
            if (!clampToAvailable) return ResolveGridMeasuredContribution(style, intrinsicContentWidth);
            double intrinsicBoxWidth = intrinsicContentWidth + style.HorizontalInsets;
            var intrinsicStyle = style.Clone();
            intrinsicStyle.ExplicitWidth = intrinsicStyle.BorderBox ? intrinsicBoxWidth : intrinsicContentWidth;
            double resolvedBoxWidth = ResolveBoxWidth(availableBoxWidth, intrinsicStyle);
            double outerWidth = style.MarginLeft + resolvedBoxWidth + style.MarginRight;
            return Math.Max(1D, clampToAvailable ? Math.Min(availableOuterWidth, outerWidth) : outerWidth);
        }
    }
}
