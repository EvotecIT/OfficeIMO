namespace OfficeIMO.Publisher.Internal;

internal sealed partial class PublisherContentsReader {
    private PublisherSourceTable? ReadTable(PublisherSourceShape source, Dictionary<uint, PublisherContentChunk> chunks) {
        uint? rows = source.Value(0x66), columns = source.Value(0x67), cellsId = source.Value(0x6B);
        PublisherBlock? dimensions = Field(source.Fields, 0x6D);
        if (!rows.HasValue || !columns.HasValue || !cellsId.HasValue || !dimensions.HasValue
            || !chunks.TryGetValue(cellsId.Value, out PublisherContentChunk? cells) || cells.Kind != 0x63) {
            _context.Add("PUB_TABLE_DEFINITION_UNRESOLVED", "A table's native grid or cell definition is unavailable.",
                OfficeConversionLossKind.Omission, PublisherEscherReader.ShapeLocation(source.Chunk.Id));
            return null;
        }
        long gridSize = (long)rows.Value * columns.Value;
        if (rows == 0 || columns == 0 || gridSize > _context.Options.Limits.MaxItems)
            throw new InvalidDataException("Publisher table grid exceeds the configured item limit or is empty.");
        var sizes = new List<double>();
        var ends = new List<uint?>();
        foreach (PublisherBlock entry in _blocks.Children(dimensions.Value)) {
            if (entry.Id != 0 || entry.Type != 0x88) continue;
            IReadOnlyList<PublisherBlock> values = _blocks.Children(entry);
            PublisherBlock size = Field(values, 2) ?? throw new InvalidDataException("Publisher table track size is missing.");
            if (size.Value == 0 || size.Value > int.MaxValue) throw new InvalidDataException("Invalid Publisher table track size.");
            sizes.Add(size.Value / 12700D);
            ends.Add(Field(values, 1)?.Value);
        }
        if (sizes.Count != rows.Value + columns.Value) throw new InvalidDataException("Publisher table track count does not match its grid.");
        var result = new PublisherSourceTable(Tracks(0, (int)columns.Value), Tracks((int)columns.Value, (int)rows.Value));
        IReadOnlyList<PublisherBlock> fields = _blocks.Chunk(cells.Offset, cells.End);
        PublisherBlock array = Field(fields, 2) ?? throw new InvalidDataException("Publisher table cells are missing.");
        var occupied = new HashSet<long>();
        foreach (PublisherBlock entry in _blocks.Children(array)) {
            if (entry.Id != 0 || entry.Type != 0x88) continue;
            IReadOnlyList<PublisherBlock> values = _blocks.Children(entry);
            // Omitted coordinates default to the first row or column.
            uint firstRow = Field(values, 1)?.Value ?? 0, lastRow = Field(values, 2)?.Value ?? 0;
            uint firstColumn = Field(values, 3)?.Value ?? 0, lastColumn = Field(values, 4)?.Value ?? 0;
            if (lastRow < firstRow || lastColumn < firstColumn || lastRow >= rows || lastColumn >= columns)
                throw new InvalidDataException("Publisher table cell span is outside its grid.");
            for (uint row = firstRow; row <= lastRow; row++) {
                for (uint column = firstColumn; column <= lastColumn; column++) {
                    _context.Record();
                    if (!occupied.Add((long)row * columns.Value + column)) throw new InvalidDataException("Overlapping Publisher table cells.");
                }
            }
            result.Cells.Add(new PublisherSourceCell((int)firstRow, (int)lastRow, (int)firstColumn, (int)lastColumn));
        }
        uint? declaredCount = Field(fields, 1)?.Value;
        if (declaredCount.HasValue && declaredCount != result.Cells.Count) throw new InvalidDataException("Publisher table cell count is inconsistent.");
        if (occupied.Count != gridSize) _context.Add("PUB_TABLE_GRID_INCOMPLETE", "Some table grid positions have no recovered cell definition.",
            OfficeConversionLossKind.Omission, PublisherEscherReader.ShapeLocation(source.Chunk.Id));
        return result;

        double[] Tracks(int start, int count) {
            bool complete = true; uint prior = 0;
            for (int i = start; i < start + count; i++) {
                _context.Token.ThrowIfCancellationRequested();
                uint? end = ends[i];
                if (!end.HasValue) { complete = false; continue; }
                if (end.Value <= prior || end.Value > int.MaxValue)
                    throw new InvalidDataException("Invalid Publisher table track endpoint order or extent.");
                prior = end.Value;
            }
            if (!complete) {
                _context.Add("PUB_TABLE_TRACK_POSITIONS_UNASSESSED", "Saved table track endpoints are incomplete. Nominal sizes are retained for this axis; laid-out cell positions are unassessed.",
                    OfficeConversionLossKind.Unassessed, PublisherEscherReader.ShapeLocation(source.Chunk.Id));
                return sizes.Skip(start).Take(count).ToArray();
            }
            // 01 stores cumulative laid-out ends; 02 retains nominal sizes even
            // when text expands a row. Each axis starts at its own zero.
            var tracks = new double[count]; prior = 0;
            for (int i = 0; i < count; i++) {
                _context.Token.ThrowIfCancellationRequested();
                uint end = ends[start + i]!.Value;
                tracks[i] = (end - prior) / 12700D; prior = end;
            }
            return tracks;
        }
    }
}

internal sealed class PublisherSourceTable {
    internal PublisherSourceTable(double[] columns, double[] rows) { Columns = columns; Rows = rows; }
    internal double[] Columns { get; }
    internal double[] Rows { get; }
    internal List<PublisherSourceCell> Cells { get; } = new();
}

internal readonly struct PublisherSourceCell {
    internal PublisherSourceCell(int firstRow, int lastRow, int firstColumn, int lastColumn) {
        FirstRow = firstRow; LastRow = lastRow; FirstColumn = firstColumn; LastColumn = lastColumn;
    }
    internal int FirstRow { get; }
    internal int LastRow { get; }
    internal int FirstColumn { get; }
    internal int LastColumn { get; }
}
