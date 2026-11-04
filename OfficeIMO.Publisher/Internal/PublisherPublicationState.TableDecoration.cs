using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Internal;

internal sealed partial class PublisherPublicationState {
    internal PublisherPublicationState SetTableCellFillColor(uint id, int row, int column, OfficeColor? color, CancellationToken token) {
        PublisherTableState table = Table(id);
        if (color.HasValue && color.Value.A != 255) throw new ArgumentException("Table fills require an opaque solid color or null.", nameof(color));
        if (row < 0 || row >= table.Rows.Count) throw new ArgumentOutOfRangeException(nameof(row));
        if (column < 0 || column >= table.Columns.Count) throw new ArgumentOutOfRangeException(nameof(column));
        int index = -1;
        for (int i = 0; i < table.Cells.Count; i++) {
            token.ThrowIfCancellationRequested();
            if (table.Cells[i].Row == row && table.Cells[i].Column == column) { index = i; break; }
        }
        if (index < 0) throw new ArgumentException("Use a cell's anchor; covered merged positions cannot be edited.", nameof(column));
        if (table.CellFills[index] == color) return this;
        OfficeColor?[] fills = table.CellFills.ToArray();
        fills[index] = color;
        return WithTableDecoration(table.WithDecoration(fills, table.BorderSegments));
    }

    internal PublisherPublicationState SetTableBorderStyle(uint id, int firstRow, int firstColumn, int lastRow, int lastColumn,
        OfficeBorderSide? side, int maximumSegments, CancellationToken token) {
        PublisherTableState table = Table(id);
        bool horizontal = firstRow == lastRow;
        if (firstRow < 0 || lastRow > table.Rows.Count || firstColumn < 0 || lastColumn > table.Columns.Count
            || (horizontal ? firstColumn >= lastColumn : firstColumn != lastColumn || firstRow >= lastRow))
            throw new ArgumentException("Select an ordered horizontal or vertical range of grid boundaries.");
        if (side.HasValue && (side.Value.Color.A != 255 || side.Value.Width <= 0 || side.Value.Width > (horizontal ? table.Height : table.Width)
            || side.Value.DashStyle != OfficeStrokeDashStyle.Solid || side.Value.LineKind != OfficeBorderLineKind.Single || side.Value.DoubleLineSeparation != 0))
            throw new ArgumentException("Table borders require an opaque solid single line with positive width within the table's perpendicular extent.", nameof(side));
        foreach (PublisherTableCellSlot cell in table.Cells) {
            token.ThrowIfCancellationRequested();
            bool interior = horizontal ? firstRow > cell.Row && firstRow < cell.Row + cell.RowSpan
                && firstColumn < cell.Column + cell.ColumnSpan && lastColumn > cell.Column
                : firstColumn > cell.Column && firstColumn < cell.Column + cell.ColumnSpan
                && firstRow < cell.Row + cell.RowSpan && lastRow > cell.Row;
            if (interior) throw new ArgumentException("The selected border range passes through a merged cell.");
        }
        int start = horizontal ? firstColumn : firstRow, end = horizontal ? lastColumn : lastRow;
        var borders = new List<PublisherTableBorderSegment>();
        bool inserted = false;
        foreach (PublisherTableBorderSegment segment in table.BorderSegments) {
            token.ThrowIfCancellationRequested();
            bool aligned = horizontal ? segment.FirstRow == segment.LastRow && segment.FirstRow == firstRow
                : segment.FirstColumn == segment.LastColumn && segment.FirstColumn == firstColumn;
            int from = horizontal ? segment.FirstColumn : segment.FirstRow, to = horizontal ? segment.LastColumn : segment.LastRow;
            if (!aligned || to <= start || from >= end) { Add(segment); continue; }
            if (from < start) Add(Range(from, start, segment.Side));
            Insert();
            if (to > end) Add(Range(end, to, segment.Side));
        }
        Insert();
        if (table.BorderSegments.SequenceEqual(borders)) return this;
        return WithTableDecoration(table.WithDecoration(table.CellFills, borders));

        PublisherTableBorderSegment Range(int from, int to, OfficeBorderSide value) => horizontal
            ? new(firstRow, from, firstRow, to, value) : new(from, firstColumn, to, firstColumn, value);
        void Insert() {
            if (!inserted && side.HasValue) Add(new(firstRow, firstColumn, lastRow, lastColumn, side.Value));
            inserted = true;
        }
        void Add(PublisherTableBorderSegment segment) {
            if (borders.Count >= maximumSegments) throw new InvalidDataException("Publisher table border segment limit exceeded.");
            borders.Add(segment);
        }
    }

    private PublisherPublicationState WithTableDecoration(PublisherTableState table) {
        var tables = Tables.ToDictionary(item => item.Key, item => item.Value);
        tables[table.Id] = table;
        return new PublisherPublicationState(Pages, ObjectOwners, Pictures, Images, Text, TableStoryReferences,
            _sourceImageIds, NextImageId, NextObjectId, Shapes, tables);
    }
}
