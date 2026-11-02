using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Internal;

/// <summary>Encodes fresh rectangular table utilities in one separately allocated OfficeArt drawing.</summary>
internal static class PublisherTableUtilityDrawingEncoding {
    internal static PublisherTableUtilityDrawing? Write(uint drawingId, IReadOnlyList<PublisherTableUtilityInput> tables,
        int maximumShapes, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        if (drawingId is 0 or > 4095 || tables.Count == 0) return null;
        int limit = Math.Min(1024, maximumShapes);
        if (limit < 3) return null;
        uint root = drawingId << 10;
        var shapes = new List<byte[]> { Root(drawingId) };
        shapes.Add(Utility(root + 1, 1, Array.Empty<(ushort, uint)>(), new[] {
            ((ushort)0x0181, 0x00FFFFFFU), ((ushort)0x01BF, 0x001F000CU), ((ushort)0x01FF, 0x00080000U)
        }));
        shapes.Add(Utility(root + 2, 20, new[] { ((ushort)0x2001, 1U) }, new[] {
            ((ushort)0x01BF, 0x00100000U), ((ushort)0x01CB, 0U), ((ushort)0x01FF, 0x00080000U)
        }));
        var ids = new HashSet<uint>();
        foreach (PublisherTableUtilityInput table in tables) {
            token.ThrowIfCancellationRequested();
            if (table.Id == 0 || !ids.Add(table.Id) || table.Rows < 1 || table.Columns < 1
                || (long)table.Rows * table.Columns != table.Fills.Count) return null;
            for (int index = 0; index < table.Fills.Count; index++) {
                token.ThrowIfCancellationRequested();
                OfficeColor? fill = table.Fills[index];
                if (shapes.Count >= limit || fill is { A: not 255 }) return null;
                uint color = fill.HasValue ? (uint)(fill.Value.R | fill.Value.G << 8 | fill.Value.B << 16) : 0x00FFFFFF;
                shapes.Add(Utility(root + (uint)shapes.Count, 1, new[] { ((ushort)0x6802, table.Id), ((ushort)0x2003, (uint)index) }, new[] {
                    ((ushort)0x0181, color), ((ushort)0x0183, 0U),
                    ((ushort)0x01BF, fill.HasValue ? 0x001F001CU : 0x001F000CU), ((ushort)0x01FF, 0x00080000U)
                }));
            }
            foreach (PublisherTableBorderSegment border in table.Borders) {
                bool horizontal = border.FirstRow == border.LastRow;
                if (border.FirstRow < 0 || border.FirstColumn < 0 || border.LastRow > table.Rows || border.LastColumn > table.Columns
                    || border.Side.Width > (horizontal ? table.Height : table.Width)
                    || (horizontal ? border.LastColumn <= border.FirstColumn : border.LastRow <= border.FirstRow || border.LastColumn != border.FirstColumn)) return null;
                int start = horizontal ? border.FirstColumn : border.FirstRow, end = horizontal ? border.LastColumn : border.LastRow;
                for (int index = start; index < end; index++) {
                    token.ThrowIfCancellationRequested();
                    if (shapes.Count >= limit) return null;
                    var unit = new PublisherTableBorderSegment(horizontal ? border.FirstRow : index, horizontal ? index : border.FirstColumn,
                        horizontal ? border.LastRow : index + 1, horizontal ? index + 1 : border.LastColumn, border.Side);
                    byte[]? encoded = PublisherTableBorderEncoding.Write(table.Id, root + (uint)shapes.Count, unit);
                    if (encoded == null) return null;
                    shapes.Add(encoded);
                }
            }
        }
        uint last = root + (uint)shapes.Count - 1;
        return new(Drawing(drawingId, shapes, last), (uint)shapes.Count, last);
    }

    internal static byte[] EmptyDrawing(uint id) => Drawing(id, new[] { Root(id) }, id << 10);

    private static byte[] Root(uint id) => Record(15, 0xF004, Bytes(writer => {
        writer.Write(Record(1, 0xF009, new byte[16]));
        writer.Write(Record(2, 0xF00A, Bytes(value => { value.Write(id << 10); value.Write(5U); })));
    }));

    private static byte[] Drawing(uint id, IEnumerable<byte[]> shapes, uint last) {
        byte[][] items = shapes.ToArray();
        return Record(15, 0xF002, Bytes(writer => {
            writer.Write(Record((ushort)(id << 4), 0xF008, Bytes(value => { value.Write((uint)items.Length); value.Write(last); })));
            writer.Write(Record(15, 0xF003, Bytes(value => { foreach (byte[] item in items) value.Write(item); })));
        }));
    }

    private static byte[] Utility(uint id, int type, (ushort, uint)[] anchor, (ushort, uint)[] properties) =>
        Record(15, 0xF004, Bytes(writer => {
            writer.Write(Record((ushort)((type << 4) | 2), 0xF00A, Bytes(value => { value.Write(id); value.Write(0xA00U); })));
            writer.Write(Record((ushort)((properties.Length << 4) | 3), 0xF00B, Bytes(value => {
                foreach (var property in properties.OrderBy(item => item.Item1)) { value.Write(property.Item1); value.Write(property.Item2); }
            })));
            var selected = anchor.Where(item => item.Item2 != 0).ToArray();
            writer.Write(Record(0, 0xF010, Bytes(value => {
                value.Write(4 + selected.Length * 6);
                foreach (var item in selected) { value.Write(item.Item1); value.Write(item.Item2); }
            })));
        }));

    private static byte[] Record(ushort initial, ushort kind, byte[] body) => Bytes(writer => {
        writer.Write(initial); writer.Write(kind); writer.Write((uint)body.Length); writer.Write(body);
    });
    private static byte[] Bytes(Action<BinaryWriter> write) {
        using var stream = new MemoryStream();
        using (var writer = new BinaryWriter(stream, System.Text.Encoding.UTF8, leaveOpen: true)) write(writer);
        return stream.ToArray();
    }
}

internal sealed record PublisherTableUtilityInput(uint Id, int Rows, int Columns, double Width, double Height, IReadOnlyList<OfficeColor?> Fills,
    IReadOnlyList<PublisherTableBorderSegment> Borders);
internal sealed record PublisherTableUtilityDrawing(byte[] Bytes, uint Count, uint LastSpid);
