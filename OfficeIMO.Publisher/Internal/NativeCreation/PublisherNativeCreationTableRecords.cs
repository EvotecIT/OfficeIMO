
namespace OfficeIMO.Publisher.Internal;

/// <summary>Encodes the current table's grid and cell ownership independently of its complete story.</summary>
internal static class PublisherNativeCreationTableRecords {
    internal static IEnumerable<byte[]> Fields(PublisherTable table, uint cellsId, PublisherParseContext context) {
        context.Token.ThrowIfCancellationRequested();
        yield return PublisherOplWriter.UInt32(0x66, (uint)table.RowCount, 0x20);
        yield return PublisherOplWriter.UInt32(0x67, (uint)table.ColumnCount, 0x20);
        yield return PublisherOplWriter.UInt32(0x68, Emu(table.Width), 0x20);
        yield return PublisherOplWriter.UInt32(0x69, Emu(table.Height), 0x20);
        yield return PublisherOplWriter.UInt32(0x6B, cellsId, 0x70);
        yield return PublisherOplWriter.Nested(0x6D, 0x90, Tracks(table.ColumnWidths, context).Concat(Tracks(table.RowHeights, context)).ToArray());
        yield return PublisherOplWriter.UInt32(0x70, unchecked((uint)-2), 0x20); // pbTableAutoFormatNone.
    }

    internal static byte[] Cells(PublisherTable table, PublisherParseContext context) {
        context.Token.ThrowIfCancellationRequested();
        ushort count = checked((ushort)table.Cells.Count);
        var cells = new byte[table.Cells.Count][];
        for (int index = 0; index < table.Cells.Count; index++) {
            context.Token.ThrowIfCancellationRequested();
            PublisherTableCell cell = table.Cells[index];
            cells[index] = PublisherOplWriter.Nested(0, 0x88,
                PublisherOplWriter.UInt32(1, (uint)cell.RowIndex, 0x20), PublisherOplWriter.UInt32(2, (uint)(cell.RowIndex + cell.RowSpan - 1), 0x20),
                PublisherOplWriter.UInt32(3, (uint)cell.ColumnIndex, 0x20), PublisherOplWriter.UInt32(4, (uint)(cell.ColumnIndex + cell.ColumnSpan - 1), 0x20),
                PublisherOplWriter.UInt32(10, 0, 0x20), PublisherOplWriter.UInt32(11, 0, 0x20), PublisherOplWriter.UInt32(12, 0, 0x20), PublisherOplWriter.UInt32(13, 0, 0x20));
        }
        byte[] result = PublisherOplWriter.Chunk(PublisherOplWriter.UInt16(1, count, 0x18), PublisherOplWriter.Nested(2, 0xA0, cells));
        context.Token.ThrowIfCancellationRequested();
        return result;
    }

    private static IEnumerable<byte[]> Tracks(IReadOnlyList<double> tracks, PublisherParseContext context) {
        context.Token.ThrowIfCancellationRequested();
        double end = 0;
        foreach (double size in tracks) {
            context.Token.ThrowIfCancellationRequested();
            end += size;
            yield return PublisherOplWriter.Nested(0, 0x88, PublisherOplWriter.UInt32(1, Emu(end), 0x20), PublisherOplWriter.UInt32(2, Emu(size), 0x20));
        }
    }

    private static uint Emu(double points) => checked((uint)(points * 12700));
}
