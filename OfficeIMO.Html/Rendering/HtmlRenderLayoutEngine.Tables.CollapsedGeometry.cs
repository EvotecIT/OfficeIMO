using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    /// <summary>Resolves authored border conflicts before measuring cells, retaining full-width winners for paint.</summary>
    private CollapsedTableBorderLayout? ResolveCollapsedTableBorderLayout(
        IElement table,
        HtmlRenderBoxStyle tableStyle,
        IReadOnlyList<TableFormattingRow> rows,
        int columnCount,
        double contentWidth,
        int legacyBorderWidth) {
        if (tableStyle.BorderCollapse != "collapse" || rows.Count == 0 || columnCount <= 0) return null;
        var winners = new Dictionary<CollapsedBorderKey, CollapsedBorderCandidate>();
        var placements = new Dictionary<TableFormattingCell, CollapsedTableCellPlacement>();
        int sourceOrder = 0;
        bool rtl = tableStyle.Direction == "rtl";
        AddHorizontalBorderRange(winners, 0, 0, columnCount, tableStyle.Borders.Top, CollapsedBorderOrigin.Table, ref sourceOrder);
        AddHorizontalBorderRange(winners, rows.Count, 0, columnCount, tableStyle.Borders.Bottom, CollapsedBorderOrigin.Table, ref sourceOrder);
        AddVerticalBorderRange(winners, rtl ? columnCount : 0, 0, rows.Count, tableStyle.Borders.Left, CollapsedBorderOrigin.Table, ref sourceOrder);
        AddVerticalBorderRange(winners, rtl ? 0 : columnCount, 0, rows.Count, tableStyle.Borders.Right, CollapsedBorderOrigin.Table, ref sourceOrder);
        AddCollapsedColumnBorders(winners, table, tableStyle, rows.Count, columnCount, contentWidth, ref sourceOrder);
        AddCollapsedRowGroupBorders(winners, rows, columnCount, rtl, ref sourceOrder);

        var occupancy = new int[columnCount];
        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
            CheckCancellation();
            TableFormattingRow row = rows[rowIndex];
            HtmlRenderBorderEdges rowBorders = row.Style.Borders;
            AddHorizontalBorderRange(winners, rowIndex, 0, columnCount, rowBorders.Top, CollapsedBorderOrigin.Row, ref sourceOrder);
            AddHorizontalBorderRange(winners, rowIndex + 1, 0, columnCount, rowBorders.Bottom, CollapsedBorderOrigin.Row, ref sourceOrder);
            AddVerticalBorderRange(winners, rtl ? columnCount : 0, rowIndex, rowIndex + 1, rowBorders.Left, CollapsedBorderOrigin.Row, ref sourceOrder);
            AddVerticalBorderRange(winners, rtl ? 0 : columnCount, rowIndex, rowIndex + 1, rowBorders.Right, CollapsedBorderOrigin.Row, ref sourceOrder);
            int column = 0;
            foreach (TableFormattingCell cell in row.Cells) {
                int requestedSpan = Math.Min(cell.ColumnSpan, columnCount);
                column = FindAvailableColumn(occupancy, column, requestedSpan);
                if (column >= columnCount) break;
                int span = Math.Max(1, Math.Min(requestedSpan, columnCount - column));
                int rowSpan = ReadRowSpan(cell.RowSpan, rows, rowIndex);
                placements.Add(cell, new CollapsedTableCellPlacement(rowIndex, column, span, rowSpan));
                HtmlRenderBoxStyle cellStyle = ResolveFormattingCellStyle(cell, contentWidth);
                ApplyTableCellFallbackInsets(cellStyle, legacyBorderWidth);
                for (int segment = column; segment < column + span; segment++) {
                    AddCollapsedBorderCandidate(winners, new CollapsedBorderKey(true, rowIndex, segment),
                        cellStyle.Borders.Top, CollapsedBorderOrigin.Cell, sourceOrder++, column, rowIndex);
                    AddCollapsedBorderCandidate(winners, new CollapsedBorderKey(true, rowIndex + rowSpan, segment),
                        cellStyle.Borders.Bottom, CollapsedBorderOrigin.Cell, sourceOrder++, column, rowIndex);
                }
                for (int segment = rowIndex; segment < rowIndex + rowSpan; segment++) {
                    AddCollapsedBorderCandidate(winners, new CollapsedBorderKey(false, rtl ? column + span : column, segment),
                        cellStyle.Borders.Left, CollapsedBorderOrigin.Cell, sourceOrder++, column, rowIndex);
                    AddCollapsedBorderCandidate(winners, new CollapsedBorderKey(false, rtl ? column : column + span, segment),
                        cellStyle.Borders.Right, CollapsedBorderOrigin.Cell, sourceOrder++, column, rowIndex);
                }
                for (int occupied = column; occupied < column + span; occupied++) {
                    occupancy[occupied] = Math.Max(occupancy[occupied], rowSpan);
                }
                column += span;
            }
            DecrementOccupancy(occupancy);
        }

        // Rows and tracks contribute only at cell edges. Their declarations
        // must not draw grid lines through a spanning cell's interior.
        foreach (CollapsedTableCellPlacement cell in placements.Values) {
            for (int boundary = cell.RowIndex + 1; boundary < cell.RowIndex + cell.RowSpan; boundary++) {
                for (int segment = cell.Column; segment < cell.Column + cell.Span; segment++) {
                    CheckCancellation();
                    ChargeLayoutOperation("collapsed-table-span-edge");
                    winners.Remove(new CollapsedBorderKey(true, boundary, segment));
                }
            }
            for (int boundary = cell.Column + 1; boundary < cell.Column + cell.Span; boundary++) {
                for (int segment = cell.RowIndex; segment < cell.RowIndex + cell.RowSpan; segment++) {
                    CheckCancellation();
                    ChargeLayoutOperation("collapsed-table-span-edge");
                    winners.Remove(new CollapsedBorderKey(false, boundary, segment));
                }
            }
        }

        var cellInsets = new Dictionary<TableFormattingCell, HtmlRenderBorderInsets>();
        foreach (KeyValuePair<TableFormattingCell, CollapsedTableCellPlacement> entry in placements) {
            CheckCancellation();
            CollapsedTableCellPlacement cell = entry.Value;
            cellInsets.Add(entry.Key, new HtmlRenderBorderInsets(
                MaximumCollapsedBorderLayoutWidth(winners, true, cell.RowIndex, cell.Column, cell.Column + cell.Span) / 2D,
                MaximumCollapsedBorderLayoutWidth(winners, false, rtl ? cell.Column : cell.Column + cell.Span, cell.RowIndex, cell.RowIndex + cell.RowSpan) / 2D,
                MaximumCollapsedBorderLayoutWidth(winners, true, cell.RowIndex + cell.RowSpan, cell.Column, cell.Column + cell.Span) / 2D,
                MaximumCollapsedBorderLayoutWidth(winners, false, rtl ? cell.Column + cell.Span : cell.Column, cell.RowIndex, cell.RowIndex + cell.RowSpan) / 2D));
        }
        return new CollapsedTableBorderLayout(winners, cellInsets, placements,
            ResolveCollapsedTablePerimeterInsets(winners, rows.Count, columnCount, rtl));
    }

    /// <summary>Keeps the table's outside half borders separate from its centerline-to-centerline cell grid.</summary>
    private static HtmlRenderBorderInsets ResolveCollapsedTablePerimeterInsets(
        IReadOnlyDictionary<CollapsedBorderKey, CollapsedBorderCandidate> winners,
        int rowCount,
        int columnCount,
        bool rtl) => new HtmlRenderBorderInsets(
            MaximumCollapsedBorderLayoutWidth(winners, true, 0, 0, columnCount) / 2D,
            MaximumCollapsedBorderLayoutWidth(winners, false, rtl ? 0 : columnCount, 0, rowCount) / 2D,
            MaximumCollapsedBorderLayoutWidth(winners, true, rowCount, 0, columnCount) / 2D,
            MaximumCollapsedBorderLayoutWidth(winners, false, rtl ? columnCount : 0, 0, rowCount) / 2D);

    private static double MaximumCollapsedBorderLayoutWidth(
        IReadOnlyDictionary<CollapsedBorderKey, CollapsedBorderCandidate> winners,
        bool horizontal,
        int boundary,
        int start,
        int end) {
        double width = 0D;
        for (int segment = start; segment < end; segment++) {
            if (winners.TryGetValue(new CollapsedBorderKey(horizontal, boundary, segment), out CollapsedBorderCandidate winner)) {
                width = Math.Max(width, winner.Border.LayoutWidth);
            }
        }
        return width;
    }

    /// <summary>A single conflict-resolution result reused by layout and paint for the supplied physical row sequence.</summary>
    private sealed class CollapsedTableBorderLayout {
        internal CollapsedTableBorderLayout(
            Dictionary<CollapsedBorderKey, CollapsedBorderCandidate> winners,
            Dictionary<TableFormattingCell, HtmlRenderBorderInsets> cellInsets,
            Dictionary<TableFormattingCell, CollapsedTableCellPlacement> placements,
            HtmlRenderBorderInsets tableInsets) {
            Winners = winners;
            CellInsets = cellInsets;
            Placements = placements;
            TableInsets = tableInsets;
        }

        internal IReadOnlyDictionary<CollapsedBorderKey, CollapsedBorderCandidate> Winners { get; }
        internal IReadOnlyDictionary<TableFormattingCell, HtmlRenderBorderInsets> CellInsets { get; }
        internal IReadOnlyDictionary<TableFormattingCell, CollapsedTableCellPlacement> Placements { get; }
        internal HtmlRenderBorderInsets TableInsets { get; }
    }

    /// <summary>Retains logical cell placement without conflating anonymous cells that share a source element.</summary>
    private readonly struct CollapsedTableCellPlacement {
        internal CollapsedTableCellPlacement(int rowIndex, int column, int span, int rowSpan) {
            RowIndex = rowIndex;
            Column = column;
            Span = span;
            RowSpan = rowSpan;
        }

        internal int RowIndex { get; }
        internal int Column { get; }
        internal int Span { get; }
        internal int RowSpan { get; }
    }
}
