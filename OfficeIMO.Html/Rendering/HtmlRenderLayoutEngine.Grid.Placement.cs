using System.Globalization;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private List<GridItem> PlaceGridItems(
        IReadOnlyList<FlexItem> items,
        int explicitColumnCount,
        HtmlRenderBoxStyle containerStyle,
        string source,
        out int columnCount,
        out int rowCount) {
        var gridItems = items
            .OrderBy(item => item.Style.Order)
            .ThenBy(item => item.SourceIndex)
            .Select(CreateGridItem)
            .ToList();
        columnCount = Math.Max(1, explicitColumnCount);
        foreach (GridItem item in gridItems) {
            if (item.RequestedColumn.HasValue) columnCount = Math.Max(columnCount, item.RequestedColumn.Value + item.ColumnSpan);
            else columnCount = Math.Max(columnCount, item.ColumnSpan);
        }
        EnsureGridPlacementLimit(columnCount);

        bool dense = containerStyle.GridAutoFlow.IndexOf("dense", StringComparison.Ordinal) >= 0;
        if (!containerStyle.GridAutoFlow.StartsWith("row", StringComparison.Ordinal)) {
            ReportUnsupportedGridValue(source, "grid-auto-flow=" + containerStyle.GridAutoFlow);
        }

        var occupied = new HashSet<long>();
        int cursorRow = 0;
        int cursorColumn = 0;
        rowCount = 0;
        foreach (GridItem item in gridItems) {
            if (item.RequestedRow.HasValue && item.RequestedColumn.HasValue) {
                item.Row = item.RequestedRow.Value;
                item.Column = item.RequestedColumn.Value;
            } else if (item.RequestedRow.HasValue) {
                item.Row = item.RequestedRow.Value;
                item.Column = FindGridColumn(occupied, item.Row, item.RowSpan, item.ColumnSpan, columnCount);
                if (item.Column + item.ColumnSpan > columnCount) columnCount = item.Column + item.ColumnSpan;
            } else if (item.RequestedColumn.HasValue) {
                item.Column = item.RequestedColumn.Value;
                item.Row = FindGridRow(occupied, item.Column, item.RowSpan, item.ColumnSpan);
            } else {
                int searchRow = dense ? 0 : cursorRow;
                int searchColumn = dense ? 0 : cursorColumn;
                FindAutomaticGridPosition(occupied, item.RowSpan, item.ColumnSpan, columnCount, ref searchRow, ref searchColumn);
                item.Row = searchRow;
                item.Column = searchColumn;
                cursorRow = searchRow;
                cursorColumn = searchColumn + item.ColumnSpan;
                if (cursorColumn >= columnCount) {
                    cursorRow++;
                    cursorColumn = 0;
                }
            }

            EnsureGridPlacementLimit(item.Row + item.RowSpan);
            EnsureGridPlacementLimit(item.Column + item.ColumnSpan);
            MarkGridArea(occupied, item.Row, item.Column, item.RowSpan, item.ColumnSpan);
            rowCount = Math.Max(rowCount, item.Row + item.RowSpan);
            columnCount = Math.Max(columnCount, item.Column + item.ColumnSpan);
        }

        rowCount = Math.Max(1, rowCount);
        return gridItems;
    }

    private GridItem CreateGridItem(FlexItem item) {
        GridAxisPlacement column = ParseGridAxisPlacement(item.Style.GridColumnStart, item.Style.GridColumnEnd, item.Source, "grid-column");
        GridAxisPlacement row = ParseGridAxisPlacement(item.Style.GridRowStart, item.Style.GridRowEnd, item.Source, "grid-row");
        return new GridItem(item, row.Start, column.Start, row.Span, column.Span);
    }

    private GridAxisPlacement ParseGridAxisPlacement(string startValue, string endValue, string source, string property) {
        GridLine start = ParseGridLine(startValue, source, property + "-start");
        GridLine end = ParseGridLine(endValue, source, property + "-end");
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

    private GridLine ParseGridLine(string value, string source, string property) {
        string normalized = string.IsNullOrWhiteSpace(value) ? "auto" : value.Trim().ToLowerInvariant();
        if (normalized == "auto") return GridLine.Auto;
        if (normalized.StartsWith("span ", StringComparison.Ordinal)
            && int.TryParse(normalized.Substring(5).Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out int span)
            && span > 0) {
            return new GridLine(GridLineKind.Span, span);
        }
        if (int.TryParse(normalized, NumberStyles.Integer, CultureInfo.InvariantCulture, out int line) && line > 0) {
            return new GridLine(GridLineKind.Line, line);
        }

        ReportUnsupportedGridValue(source, property + "=" + value);
        return GridLine.Auto;
    }

    private static int FindGridColumn(HashSet<long> occupied, int row, int rowSpan, int columnSpan, int columnCount) {
        for (int column = 0; column + columnSpan <= columnCount; column++) {
            if (CanPlaceGridArea(occupied, row, column, rowSpan, columnSpan)) return column;
        }
        return columnCount;
    }

    private static int FindGridRow(HashSet<long> occupied, int column, int rowSpan, int columnSpan) {
        for (int row = 0; ; row++) {
            if (CanPlaceGridArea(occupied, row, column, rowSpan, columnSpan)) return row;
        }
    }

    private static void FindAutomaticGridPosition(HashSet<long> occupied, int rowSpan, int columnSpan, int columnCount, ref int row, ref int column) {
        for (;;) {
            if (column + columnSpan <= columnCount && CanPlaceGridArea(occupied, row, column, rowSpan, columnSpan)) return;
            column++;
            if (column + columnSpan > columnCount) {
                row++;
                column = 0;
            }
        }
    }

    private static bool CanPlaceGridArea(HashSet<long> occupied, int row, int column, int rowSpan, int columnSpan) {
        for (int rowOffset = 0; rowOffset < rowSpan; rowOffset++) {
            for (int columnOffset = 0; columnOffset < columnSpan; columnOffset++) {
                if (occupied.Contains(GridCellKey(row + rowOffset, column + columnOffset))) return false;
            }
        }
        return true;
    }

    private static void MarkGridArea(HashSet<long> occupied, int row, int column, int rowSpan, int columnSpan) {
        for (int rowOffset = 0; rowOffset < rowSpan; rowOffset++) {
            for (int columnOffset = 0; columnOffset < columnSpan; columnOffset++) {
                occupied.Add(GridCellKey(row + rowOffset, column + columnOffset));
            }
        }
    }

    private static long GridCellKey(int row, int column) => ((long)row << 32) | (uint)column;

    private void EnsureGridPlacementLimit(int count) {
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
