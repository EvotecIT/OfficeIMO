using static OfficeIMO.Publisher.WriterProbe.NativeEncoding;

namespace OfficeIMO.Publisher.WriterProbe;

/// <summary>Encodes the current table's grid and cell ownership independently of its complete story.</summary>
internal static class CandidateTableRecords {
    internal static IEnumerable<byte[]> Fields(PublisherTable table, uint cellsId) {
        yield return Scalar(0x66, 0x20, (uint)table.RowCount);
        yield return Scalar(0x67, 0x20, (uint)table.ColumnCount);
        yield return Scalar(0x68, 0x20, Emu(table.Width));
        yield return Scalar(0x69, 0x20, Emu(table.Height));
        yield return Scalar(0x6B, 0x70, cellsId);
        yield return Variable(0x6D, 0x90, Tracks(table.ColumnWidths).Concat(Tracks(table.RowHeights)).ToArray());
        yield return Scalar(0x70, 0x20, unchecked((uint)-2)); // pbTableAutoFormatNone.
    }

    internal static byte[] Cells(PublisherTable table) => Chunk(ShortScalar(1, 0x18, checked((ushort)table.Cells.Count)),
        Variable(2, 0xA0, table.Cells.Select(cell => Variable(0, 0x88,
            Scalar(1, 0x20, (uint)cell.RowIndex), Scalar(2, 0x20, (uint)(cell.RowIndex + cell.RowSpan - 1)),
            Scalar(3, 0x20, (uint)cell.ColumnIndex), Scalar(4, 0x20, (uint)(cell.ColumnIndex + cell.ColumnSpan - 1)),
            Scalar(10, 0x20, 0), Scalar(11, 0x20, 0), Scalar(12, 0x20, 0), Scalar(13, 0x20, 0))).ToArray()));

    private static IEnumerable<byte[]> Tracks(IReadOnlyList<double> tracks) {
        double end = 0;
        foreach (double size in tracks) {
            end += size;
            yield return Variable(0, 0x88, Scalar(1, 0x20, Emu(end)), Scalar(2, 0x20, Emu(size)));
        }
    }

    private static uint Emu(double points) => checked((uint)(points * 12700));
}
