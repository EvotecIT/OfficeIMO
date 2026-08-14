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

    private double ResolveInlineGridWidth(IElement element, double availableWidth, HtmlRenderBoxStyle style, int depth) {
        double availableOuterWidth = Math.Max(1D, availableWidth);
        double availableBoxWidth = Math.Max(1D, availableOuterWidth - style.MarginLeft - style.MarginRight);
        if (style.ExplicitWidth.HasValue) {
            return Math.Min(availableOuterWidth, style.MarginLeft + ResolveBoxWidth(availableBoxWidth, style) + style.MarginRight);
        }

        if (!TryCollectFlexItems(element, availableOuterWidth, style, depth, out List<FlexItem> formattingItems)) return availableOuterWidth;
        string source = HtmlRenderStyleResolver.DescribeSource(element);
        List<GridTrack> tracks = ParseGridTracks(style.GridTemplateColumns, availableBoxWidth, percentageReferenceIsDefinite: true, style, source, "grid-template-columns");
        double? declaredContentHeight = ResolveGridDeclaredContentHeight(style);
        List<GridTrack> rows = ParseGridTracks(style.GridTemplateRows, declaredContentHeight ?? 0D, declaredContentHeight.HasValue, style, source, "grid-template-rows");
        IReadOnlyDictionary<string, GridAreaDefinition> areas = ParseGridTemplateAreas(style.GridTemplateAreas, source, out int areaRowCount, out int areaColumnCount);
        IReadOnlyDictionary<string, int> columnLineNames = ParseGridLineNames(style.GridTemplateColumns);
        IReadOnlyDictionary<string, int> rowLineNames = ParseGridLineNames(style.GridTemplateRows);
        int explicitColumns = Math.Max(1, Math.Max(tracks.Count, areaColumnCount));
        int explicitRows = Math.Max(1, Math.Max(rows.Count, areaRowCount));
        List<GridItem> items = PlaceGridItems(formattingItems, explicitColumns, explicitRows, style, source, areas, columnLineNames, rowLineNames, out int columnCount, out _);
        CollapseEmptyAutoFitColumns(style, items, tracks, ref columnCount);
        EnsureGridTrackCount(tracks, columnCount, style.GridAutoColumns, availableBoxWidth, percentageReferenceIsDefinite: true, style, source, "grid-auto-columns");
        List<double> sizes = ResolveGridIntrinsicTrackBases(
            tracks,
            items,
            availableBoxWidth,
            style.ColumnGap,
            includeFractionTracks: true);

        double intrinsicContentWidth = sizes.Sum() + style.ColumnGap * CountGridBaseGaps(tracks);
        double intrinsicBoxWidth = intrinsicContentWidth + style.HorizontalInsets;
        var intrinsicStyle = style.Clone();
        intrinsicStyle.ExplicitWidth = intrinsicStyle.BorderBox ? intrinsicBoxWidth : intrinsicContentWidth;
        double resolvedBoxWidth = ResolveBoxWidth(availableBoxWidth, intrinsicStyle);
        return Math.Max(1D, Math.Min(availableOuterWidth, style.MarginLeft + resolvedBoxWidth + style.MarginRight));
    }
}
