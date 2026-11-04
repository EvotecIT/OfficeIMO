namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private List<GridIntrinsicContribution> CollectGridIntrinsicContributions(
        IReadOnlyList<GridItem> items,
        double availableSize,
        double parentColumnGap,
        IReadOnlyDictionary<string, int> columnLineNames,
        IReadOnlyDictionary<string, int> rowLineNames,
        int depth) {
        var contributions = new List<GridIntrinsicContribution>();
        foreach (GridItem item in items) {
            CheckCancellation();
            HtmlRenderBoxStyle style = item.Item.Style;
            if (item.Item.Element == null || style.Display is not ("grid" or "inline-grid") || !IsSubgridTrackList(style.GridTemplateColumns)
                || !TryCollectFlexItems(item.Item.Element, availableSize, style, depth + 1, captureRunningElements: false,
                    out List<FlexItem> children, out _, registerOutOfFlowElements: false) || children.Count == 0) {
                contributions.Add(new GridIntrinsicContribution(item, 0D));
                continue;
            }

            // A column subgrid contributes its children to the parent tracks that
            // they occupy. Measuring all of its text as one spanning item inflates
            // unrelated auto columns and can starve a minmax(0,1fr) title column.
            string source = item.Item.Source;
            bool rowSubgrid = IsSubgridTrackList(style.GridTemplateRows);
            double? declaredHeight = ResolveGridDeclaredContentHeight(style);
            int explicitRows = rowSubgrid ? item.RowSpan : ParseGridTracks(style.GridTemplateRows, declaredHeight ?? 0D, declaredHeight.HasValue, style, source, "grid-template-rows").Count;
            IReadOnlyDictionary<string, GridAreaDefinition> areas = ParseGridTemplateAreas(style.GridTemplateAreas, source, out int areaRows, out _);
            var inheritedColumns = columnLineNames.Where(pair => pair.Value >= item.Column && pair.Value <= item.Column + item.ColumnSpan)
                .ToDictionary(pair => pair.Key, pair => pair.Value - item.Column, StringComparer.Ordinal);
            var inheritedRows = rowSubgrid ? rowLineNames.Where(pair => pair.Value >= item.Row && pair.Value <= item.Row + item.RowSpan)
                .ToDictionary(pair => pair.Key, pair => pair.Value - item.Row, StringComparer.Ordinal) : null;
            IReadOnlyDictionary<string, int> childColumns = AddGridAreaLineNames(
                ParseGridLineNames(style.GridTemplateColumns, item.ColumnSpan + 1, inheritedColumns), areas, rows: false);
            IReadOnlyDictionary<string, int> childRows = AddGridAreaLineNames(
                ParseGridLineNames(style.GridTemplateRows, rowSubgrid ? item.RowSpan + 1 : null, inheritedRows), areas, rows: true);
            List<GridItem> placed = PlaceGridItems(children, item.ColumnSpan, rowSubgrid ? explicitRows : Math.Max(explicitRows, areaRows),
                style, source, childColumns, childRows, out _, out _, out _, out int leadingRows,
                allowImplicitColumns: false, allowImplicitRows: !rowSubgrid);
            ClampSubgridPlacements(placed, item.ColumnSpan, rows: false);
            childRows = OffsetGridLineNames(childRows, leadingRows);
            double childColumnGap = style.ColumnGapWasSpecified ? style.ColumnGap : parentColumnGap;
            double innerEdgeInset = (childColumnGap - parentColumnGap) / 2D;
            List<GridIntrinsicContribution> descendants = CollectGridIntrinsicContributions(placed, availableSize, childColumnGap,
                childColumns, childRows, depth + 1);
            foreach (GridIntrinsicContribution descendant in descendants) {
                GridItem child = descendant.Item;
                double edgeInsets = descendant.EdgeInsets;
                if (child.Column == 0) edgeInsets += style.MarginLeft + style.BorderLeftWidth + style.PaddingLeft;
                else edgeInsets += innerEdgeInset;
                if (child.Column + child.ColumnSpan == item.ColumnSpan) edgeInsets += style.MarginRight + style.BorderRightWidth + style.PaddingRight;
                else edgeInsets += innerEdgeInset;
                child.Column += item.Column;
                child.Row += item.Row;
                contributions.Add(new GridIntrinsicContribution(child, edgeInsets));
            }
        }
        return contributions;
    }

    /// <summary>
    /// A row subgrid's children size the parent rows they occupy. Measuring the subgrid as one
    /// spanning item would share its height evenly and leave the tallest row too short.
    /// </summary>
    private List<double> ResolveGridRowsWithSubgridContributions(
        List<GridTrack> rowTracks,
        List<GridItem> items,
        double rowGap,
        double? declaredContentHeight) {
        var measured = new List<(GridItem Item, double[] Rows)>();
        foreach (GridItem item in items) {
            if (item.Item.Element != null && _measuredRowSubgridRows.TryGetValue(item.Item.Element, out double[]? rows)) {
                _measuredRowSubgridRows.Remove(item.Item.Element);
                measured.Add((item, rows));
            }
        }
        if (measured.Count == 0) return ResolveNaturalGridRows(rowTracks, items, rowGap, declaredContentHeight);

        var ordinary = items.Where(item => !measured.Any(entry => ReferenceEquals(entry.Item, item))).ToList();
        List<double> sizes = ResolveNaturalGridRows(rowTracks, ordinary, rowGap, declaredContentHeight);
        foreach ((GridItem item, double[] rows) in measured) {
            for (int index = 0; index < rows.Length && index < item.RowSpan && item.Row + index < sizes.Count; index++) {
                int row = item.Row + index;
                if (rowTracks[row].Kind != GridTrackKind.Fixed) sizes[row] = Math.Max(sizes[row], rows[index]);
            }
        }
        return sizes;
    }

    private static double[] MeasureRowSubgridContribution(IReadOnlyList<double> rowSizes, HtmlRenderBoxStyle style) {
        double[] rows = rowSizes.ToArray();
        if (rows.Length == 0) return rows;
        rows[0] += style.MarginTop + style.BorderTopWidth + style.PaddingTop;
        rows[rows.Length - 1] += style.MarginBottom + style.BorderBottomWidth + style.PaddingBottom;
        return rows;
    }

    private sealed class GridIntrinsicContribution {
        internal GridIntrinsicContribution(GridItem item, double edgeInsets) {
            Item = item;
            EdgeInsets = edgeInsets;
        }

        internal GridItem Item { get; }
        internal double EdgeInsets { get; }
    }
}
