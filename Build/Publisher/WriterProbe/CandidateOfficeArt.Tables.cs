using OfficeIMO.Drawing;
using OfficeIMO.Publisher.Internal;
using static OfficeIMO.Publisher.WriterProbe.NativeEncoding;
using static OfficeIMO.Publisher.WriterProbe.NativePublicationGraph;

namespace OfficeIMO.Publisher.WriterProbe;

internal static partial class CandidateOfficeArt {
    /// <summary>Allocates fresh table utility records in the separately owned third drawing, without reading a native publication.</summary>
    private static (byte[] Bytes, uint Count, uint LastSpid)? TableDrawing(Graph graph) {
        ObjectBinding[] tables = graph.Objects.Where(item => item.Source.Table is { } table
            && (table.Cells.Any(cell => cell.FillColor.HasValue) || table.BorderSegments.Count != 0)).ToArray();
        if (tables.Length == 0) return null;
        const uint id = 3, root = id << 10;
        var shapes = new List<byte[]> { DrawingRoot(id) };
        // Native utility drawings carry default cell and border styles before
        // table-owned records. Explicit no-fill/no-line defaults keep untouched cells borderless.
        shapes.Add(Utility(root + 1, 1, Array.Empty<(ushort, uint)>(), new[] {
            ((ushort)0x0181, 0x00FFFFFFU), ((ushort)0x01BF, 0x001F000CU), ((ushort)0x01FF, 0x00080000U)
        }));
        shapes.Add(Utility(root + 2, 20, new[] { ((ushort)0x2001, 1U) }, new[] {
            ((ushort)0x01BF, 0x00100000U), ((ushort)0x01CB, 0U), ((ushort)0x01FF, 0x00080000U)
        }));
        foreach (ObjectBinding item in tables) {
            PublisherTable table = item.Source.Table!;
            if (table.Cells.Any(cell => cell.RowSpan != 1 || cell.ColumnSpan != 1))
                throw new NotSupportedException("Decorated merged-table native creation requires separate qualification.");
            for (int index = 0; index < table.Cells.Count; index++) {
                OfficeColor? fill = table.Cells[index].FillColor;
                if (fill is { A: not 255 }) throw new NotSupportedException("The table utility codec supports opaque solid colors.");
                Add(new[] { ((ushort)0x6802, item.Id), ((ushort)0x2003, (uint)index) }, new[] {
                    ((ushort)0x0181, fill.HasValue ? Color(fill.Value) : 0x00FFFFFFU), ((ushort)0x0183, 0U),
                    ((ushort)0x01BF, fill.HasValue ? 0x001F001CU : 0x001F000CU), ((ushort)0x01FF, 0x00080000U)
                });
            }
            foreach (PublisherTableBorderSegment border in table.BorderSegments) {
                if (shapes.Count >= 1024) throw new NotSupportedException("The table utility specimen is bounded to one SPID cluster.");
                shapes.Add(PublisherTableBorderEncoding.Write(item.Id, root + (uint)shapes.Count, border)
                    ?? throw new NotSupportedException("The table utility codec requires opaque solid single borders with exact positive signed-EMU widths."));
            }
        }
        uint last = root + (uint)shapes.Count - 1;
        return (Drawing(id, shapes.ToArray(), last), (uint)shapes.Count, last);

        void Add((ushort Id, uint Value)[] anchor, (ushort Id, uint Value)[] properties) {
            if (shapes.Count >= 1024) throw new NotSupportedException("The table utility specimen is bounded to one SPID cluster.");
            shapes.Add(Utility(root + (uint)shapes.Count, 1, anchor, properties));
        }
    }

    private static uint Color(OfficeColor color) => (uint)(color.R | color.G << 8 | color.B << 16);

    private static byte[] Utility(uint spid, int type, (ushort Id, uint Value)[] anchor, (ushort Id, uint Value)[] properties) =>
        Art(15, 0xF004, Art(checked((ushort)((type << 4) | 2)), 0xF00A, Bytes(writer => { writer.Write(spid); writer.Write(0xA00U); })),
            Art(checked((ushort)((properties.Length << 4) | 3)), 0xF00B, Bytes(writer => {
                foreach (var property in properties.OrderBy(item => item.Id)) { writer.Write(property.Id); writer.Write(property.Value); }
            })), Art(0, 0xF010, Client(anchor.Where(item => item.Value != 0).ToArray())));
}
