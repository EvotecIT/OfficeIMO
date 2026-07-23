using System.Globalization;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private List<GridItem> PlaceGridItems(
        IReadOnlyList<FlexItem> items,
        int explicitColumnCount,
        int explicitRowCount,
        HtmlRenderBoxStyle containerStyle,
        string source,
        IReadOnlyDictionary<string, GridAreaDefinition> areas,
        IReadOnlyDictionary<string, int> columnLineNames,
        IReadOnlyDictionary<string, int> rowLineNames,
        out int columnCount,
        out int rowCount) {
        var gridItems = items
            .OrderBy(item => item.Style.Order)
            .ThenBy(item => item.SourceIndex)
            .Select(item => CreateGridItem(item, areas, columnLineNames, rowLineNames))
            .ToList();
        columnCount = Math.Max(1, explicitColumnCount);
        foreach (GridItem item in gridItems) {
            if (item.RequestedColumn.HasValue) columnCount = Math.Max(columnCount, GridPlacementEnd(item.RequestedColumn.Value, item.ColumnSpan));
            else columnCount = Math.Max(columnCount, item.ColumnSpan);
        }
        EnsureGridPlacementLimit(columnCount);

        IReadOnlyList<string> autoFlowTokens = HtmlRenderCssValues.SplitWhitespace(containerStyle.GridAutoFlow);
        bool dense = autoFlowTokens.Contains("dense");
        bool columnFlow = autoFlowTokens.Contains("column");
        if (autoFlowTokens.Any(token => token != "row" && token != "column" && token != "dense") || autoFlowTokens.Contains("row") && columnFlow) {
            ReportUnsupportedGridValue(source, "grid-auto-flow=" + containerStyle.GridAutoFlow);
            columnFlow = false;
        }

        var occupied = new HashSet<long>();
        int cursorRow = 0;
        int cursorColumn = 0;
        rowCount = Math.Max(1, explicitRowCount);
        foreach (GridItem item in gridItems) {
            if (item.RequestedRow.HasValue && item.RequestedColumn.HasValue) {
                item.Row = item.RequestedRow.Value;
                item.Column = item.RequestedColumn.Value;
            } else if (item.RequestedRow.HasValue) {
                item.Row = item.RequestedRow.Value;
                item.Column = FindGridColumn(occupied, item.Row, item.RowSpan, item.ColumnSpan, columnCount);
                int requestedColumnEnd = GridPlacementEnd(item.Column, item.ColumnSpan);
                if (requestedColumnEnd > columnCount) columnCount = requestedColumnEnd;
            } else if (item.RequestedColumn.HasValue) {
                item.Column = item.RequestedColumn.Value;
                item.Row = FindGridRow(occupied, item.Column, item.RowSpan, item.ColumnSpan);
            } else {
                int searchRow = dense ? 0 : cursorRow;
                int searchColumn = dense ? 0 : cursorColumn;
                int columnFlowRowCount = Math.Max(explicitRowCount, item.RowSpan);
                EnsureGridPlacementLimit(columnFlowRowCount);
                if (columnFlow) FindAutomaticGridPositionColumn(occupied, item.RowSpan, item.ColumnSpan, columnFlowRowCount, ref searchRow, ref searchColumn);
                else FindAutomaticGridPosition(occupied, item.RowSpan, item.ColumnSpan, columnCount, ref searchRow, ref searchColumn);
                item.Row = searchRow;
                item.Column = searchColumn;
                if (columnFlow) {
                    cursorRow = GridPlacementEnd(searchRow, item.RowSpan);
                    cursorColumn = searchColumn;
                    if (cursorRow >= columnFlowRowCount) {
                        cursorColumn++;
                        cursorRow = 0;
                    }
                } else {
                    cursorRow = searchRow;
                    cursorColumn = GridPlacementEnd(searchColumn, item.ColumnSpan);
                    if (cursorColumn >= columnCount) {
                        cursorRow++;
                        cursorColumn = 0;
                    }
                }
            }

            int itemRowEnd = GridPlacementEnd(item.Row, item.RowSpan);
            int itemColumnEnd = GridPlacementEnd(item.Column, item.ColumnSpan);
            MarkGridArea(occupied, item.Row, item.Column, item.RowSpan, item.ColumnSpan);
            rowCount = Math.Max(rowCount, itemRowEnd);
            columnCount = Math.Max(columnCount, itemColumnEnd);
        }

        rowCount = Math.Max(1, rowCount);
        return gridItems;
    }

    private GridItem CreateGridItem(
        FlexItem item,
        IReadOnlyDictionary<string, GridAreaDefinition> areas,
        IReadOnlyDictionary<string, int> columnLineNames,
        IReadOnlyDictionary<string, int> rowLineNames) {
        string areaName = item.Style.GridArea;
        if (areaName != "auto" && areaName.IndexOf('/') < 0 && !int.TryParse(areaName, out _) && !areaName.StartsWith("span ", StringComparison.Ordinal)) {
            if (areas.TryGetValue(areaName, out GridAreaDefinition? area)) {
                return new GridItem(item, area.Row, area.Column, area.RowSpan, area.ColumnSpan);
            }
            ReportUnsupportedGridValue(item.Source, "grid-area=" + areaName);
            return new GridItem(item, null, null, 1, 1);
        }
        GridAxisPlacement column = ParseGridAxisPlacement(item.Style.GridColumnStart, item.Style.GridColumnEnd, item.Source, "grid-column", columnLineNames);
        GridAxisPlacement row = ParseGridAxisPlacement(item.Style.GridRowStart, item.Style.GridRowEnd, item.Source, "grid-row", rowLineNames);
        return new GridItem(item, row.Start, column.Start, row.Span, column.Span);
    }

    private GridAxisPlacement ParseGridAxisPlacement(
        string startValue,
        string endValue,
        string source,
        string property,
        IReadOnlyDictionary<string, int> lineNames) {
        GridLine start = ParseGridLine(startValue, source, property + "-start", lineNames);
        GridLine end = ParseGridLine(endValue, source, property + "-end", lineNames);
        int span = start.Kind == GridLineKind.Span ? start.Value : end.Kind == GridLineKind.Span ? end.Value : 1;
        int? position = null;
        if (start.Kind == GridLineKind.Line) {
            position = start.Value - 1;
            if (end.Kind == GridLineKind.Line && end.Value > start.Value) span = end.Value - start.Value;
        } else if (end.Kind == GridLineKind.Line) {
            position = Math.Max(0, end.Value - 1 - span);
        }

        return new GridAxisPlacement(position, Math.Max(1, span));
    }

    private GridLine ParseGridLine(string value, string source, string property, IReadOnlyDictionary<string, int> lineNames) {
        string normalized = string.IsNullOrWhiteSpace(value) ? "auto" : value.Trim().ToLowerInvariant();
        if (normalized == "auto") return GridLine.Auto;
        if (normalized.StartsWith("span ", StringComparison.Ordinal)
            && int.TryParse(normalized.Substring(5).Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out int span)
            && span > 0) {
            EnsureGridPlacementLimit(span);
            return new GridLine(GridLineKind.Span, span);
        }
        if (int.TryParse(normalized, NumberStyles.Integer, CultureInfo.InvariantCulture, out int line) && line > 0) {
            EnsureGridPlacementLimit(line);
            return new GridLine(GridLineKind.Line, line);
        }
        if (lineNames.TryGetValue(normalized, out int namedLine)) return new GridLine(GridLineKind.Line, namedLine + 1);

        ReportUnsupportedGridValue(source, property + "=" + value);
        return GridLine.Auto;
    }

    private int FindGridColumn(HashSet<long> occupied, int row, int rowSpan, int columnSpan, int columnCount) {
        for (int column = 0; column + columnSpan <= columnCount; column++) {
            if (CanPlaceGridArea(occupied, row, column, rowSpan, columnSpan)) return column;
        }
        return columnCount;
    }

    private int FindGridRow(HashSet<long> occupied, int column, int rowSpan, int columnSpan) {
        for (int row = 0; ; row++) {
            if (CanPlaceGridArea(occupied, row, column, rowSpan, columnSpan)) return row;
        }
    }

    private void FindAutomaticGridPosition(HashSet<long> occupied, int rowSpan, int columnSpan, int columnCount, ref int row, ref int column) {
        for (;;) {
            if (column + columnSpan <= columnCount && CanPlaceGridArea(occupied, row, column, rowSpan, columnSpan)) return;
            column++;
            if (column + columnSpan > columnCount) {
                row++;
                column = 0;
            }
        }
    }

    private void FindAutomaticGridPositionColumn(HashSet<long> occupied, int rowSpan, int columnSpan, int rowCount, ref int row, ref int column) {
        for (;;) {
            if (row + rowSpan <= rowCount && CanPlaceGridArea(occupied, row, column, rowSpan, columnSpan)) return;
            row++;
            if (row + rowSpan > rowCount) {
                column++;
                row = 0;
            }
        }
    }

    private bool CanPlaceGridArea(HashSet<long> occupied, int row, int column, int rowSpan, int columnSpan) {
        for (int rowOffset = 0; rowOffset < rowSpan; rowOffset++) {
            for (int columnOffset = 0; columnOffset < columnSpan; columnOffset++) {
                ChargeLayoutOperation("CSS grid placement");
                if (occupied.Contains(GridCellKey(row + rowOffset, column + columnOffset))) return false;
            }
        }
        return true;
    }

    private void MarkGridArea(HashSet<long> occupied, int row, int column, int rowSpan, int columnSpan) {
        for (int rowOffset = 0; rowOffset < rowSpan; rowOffset++) {
            for (int columnOffset = 0; columnOffset < columnSpan; columnOffset++) {
                ChargeLayoutOperation("CSS grid placement");
                occupied.Add(GridCellKey(row + rowOffset, column + columnOffset));
            }
        }
    }

    private static long GridCellKey(int row, int column) => ((long)row << 32) | (uint)column;

    private int GridPlacementEnd(int start, int span) {
        long end = (long)start + span;
        EnsureGridPlacementLimit(end);
        return (int)end;
    }

    private void EnsureGridPlacementLimit(long count) {
        if (count <= _options.MaxGridTracks) return;
        throw new HtmlDomLimitException(
            HtmlRenderDiagnosticCodes.GridTrackLimitExceeded,
            "Grid placement exceeded the configured maximum track count.",
            nameof(HtmlRenderOptions.MaxGridTracks),
            count,
            _options.MaxGridTracks);
    }

    private enum GridLineKind {
        Auto,
        Line,
        Span
    }

    private readonly struct GridLine {
        internal GridLine(GridLineKind kind, int value) {
            Kind = kind;
            Value = value;
        }
        internal static GridLine Auto => new GridLine(GridLineKind.Auto, 0);
        internal GridLineKind Kind { get; }
        internal int Value { get; }
    }

    private readonly struct GridAxisPlacement {
        internal GridAxisPlacement(int? start, int span) {
            Start = start;
            Span = span;
        }
        internal int? Start { get; }
        internal int Span { get; }
    }

    private sealed class GridItem {
        internal GridItem(FlexItem item, int? requestedRow, int? requestedColumn, int rowSpan, int columnSpan) {
            Item = item;
            RequestedRow = requestedRow;
            RequestedColumn = requestedColumn;
            RowSpan = rowSpan;
            ColumnSpan = columnSpan;
            HasExplicitWidth = item.Style.ExplicitWidth.HasValue;
            HasExplicitHeight = item.Style.ExplicitHeight.HasValue;
        }
        internal FlexItem Item { get; }
        internal int? RequestedRow { get; }
        internal int? RequestedColumn { get; }
        internal int RowSpan { get; }
        internal int ColumnSpan { get; }
        internal int Row { get; set; }
        internal int Column { get; set; }
        internal bool HasExplicitWidth { get; }
        internal bool HasExplicitHeight { get; }
        internal HtmlRenderFlowBlock? Block { get; set; }
        internal double OffsetX { get; set; }
        internal double OffsetY { get; set; }
    }
}
