using OfficeIMO.Drawing;
using OfficeIMO.Drawing.Binary;

namespace OfficeIMO.Publisher.Internal;

/// <summary>Encodes a fresh solid shared-edge utility, without publication-object identity or source-template bytes.</summary>
internal static class PublisherTableBorderEncoding {
    internal static byte[]? Write(uint tableId, uint spid, PublisherTableBorderSegment border) {
        OfficeBorderSide side = border.Side;
        double width = Math.Round(side.Width * 12700);
        if (tableId == 0 || spid == 0 || side.Color.A != 255 || side.Width <= 0 || width > int.MaxValue || side.Width != width / 12700
            || side.DashStyle != OfficeStrokeDashStyle.Solid || side.LineKind != OfficeBorderLineKind.Single || side.DoubleLineSeparation != 0) return null;
        var values = new Dictionary<ushort, uint> {
            [0x0181] = (uint)(side.Color.R | side.Color.G << 8 | side.Color.B << 16), [0x0183] = uint.MaxValue,
            [0x01BF] = 0x001F001C, [0x01CB] = (uint)width, [0x01FF] = 0x00080000
        };
        byte[]? properties = OfficeArtPropertyTableEditor.RewriteScalars(Array.Empty<byte>(), 0, 0, 0, values, out ushort count);
        if (properties == null) return null;
        var anchor = new[] { ((ushort)0x2001, border.FirstRow == border.LastRow ? 1U : 2U), ((ushort)0x6802, tableId),
            ((ushort)0x2004, (uint)border.FirstRow), ((ushort)0x2005, (uint)border.FirstColumn),
            ((ushort)0x2006, (uint)border.LastRow), ((ushort)0x2007, (uint)border.LastColumn) }.Where(item => item.Item2 != 0).ToArray();
        byte[] fsp = Record(0x12, 0xF00A, Bytes(writer => { writer.Write(spid); writer.Write(0xA00U); }));
        byte[] fopt = Record((ushort)((count << 4) | 3), 0xF00B, properties);
        byte[] client = Record(0, 0xF010, Bytes(writer => {
            writer.Write(4 + anchor.Length * 6);
            foreach (var value in anchor) { writer.Write(value.Item1); writer.Write(value.Item2); }
        }));
        return Record(15, 0xF004, Bytes(writer => { writer.Write(fsp); writer.Write(fopt); writer.Write(client); }));
    }

    private static byte[] Record(ushort initial, ushort kind, byte[] body) => Bytes(writer => {
        writer.Write(initial); writer.Write(kind); writer.Write(body.Length); writer.Write(body);
    });
    private static byte[] Bytes(Action<BinaryWriter> write) {
        using var stream = new MemoryStream();
        using (var writer = new BinaryWriter(stream, System.Text.Encoding.UTF8, leaveOpen: true)) write(writer);
        return stream.ToArray();
    }
}
