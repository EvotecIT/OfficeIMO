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
        OfficeBorderSide side, CancellationToken token) {
        PublisherTableState table = Table(id);
        bool horizontal = firstRow == lastRow;
        if (firstRow < 0 || lastRow > table.Rows.Count || firstColumn < 0 || lastColumn > table.Columns.Count
            || (horizontal ? firstColumn >= lastColumn : firstColumn != lastColumn || firstRow >= lastRow))
            throw new ArgumentException("Select an ordered horizontal or vertical range of grid boundaries.");
        if (side.Color.A != 255 || side.Width <= 0 || side.Width > (horizontal ? table.Height : table.Width)
            || side.DashStyle != OfficeStrokeDashStyle.Solid || side.LineKind != OfficeBorderLineKind.Single || side.DoubleLineSeparation != 0)
            throw new ArgumentException("Table borders require an opaque solid single line with positive width within the table's perpendicular extent.", nameof(side));
        int index = -1;
        for (int i = 0; i < table.BorderSegments.Count; i++) {
            token.ThrowIfCancellationRequested();
            PublisherTableBorderSegment segment = table.BorderSegments[i];
            if (segment.FirstRow != firstRow || segment.FirstColumn != firstColumn || segment.LastRow != lastRow || segment.LastColumn != lastColumn) continue;
            if (index >= 0) throw new NotSupportedException("The selected border has ambiguous shared ownership.");
            index = i;
        }
        if (index < 0) throw new ArgumentException("Select the exact endpoints of one existing shared border.");
        if (table.BorderSegments[index].Side == side) return this;
        PublisherTableBorderSegment[] borders = table.BorderSegments.ToArray();
        borders[index] = new(firstRow, firstColumn, lastRow, lastColumn, side);
        return WithTableDecoration(table.WithDecoration(table.CellFills, borders));
    }

    private PublisherPublicationState WithTableDecoration(PublisherTableState table) {
        var tables = Tables.ToDictionary(item => item.Key, item => item.Value);
        tables[table.Id] = table;
        return new PublisherPublicationState(Pages, ObjectOwners, Pictures, Images, Text, TableStoryReferences,
            _sourceImageIds, NextImageId, NextObjectId, Shapes, tables);
    }
}
