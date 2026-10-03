using System.Buffers.Binary;
using System.Globalization;
using System.IO.Compression;
using System.Text;
using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Benchmarks;

// Synthetic supported-format input, not native Apple qualification or a product writer.
internal static class IWorkScaleInput {
    internal static string Text(int index) => "Item " + index.ToString(CultureInfo.InvariantCulture);

    internal static byte[] Create(IWorkDocumentKind kind, int units) {
        if (units < 1 || units > 100_000) throw new ArgumentOutOfRangeException(nameof(units));
        byte[] records = kind switch {
            IWorkDocumentKind.Pages => Pages(units),
            IWorkDocumentKind.Numbers => Numbers(units),
            IWorkDocumentKind.Keynote => Keynote(units),
            _ => throw new ArgumentOutOfRangeException(nameof(kind))
        };
        byte[] snappy = Join(Varint((ulong)records.Length), LiteralHeader(records.Length), records);
        if (snappy.Length > 0xffffff) throw new InvalidDataException("Synthetic IWA chunk exceeds its 24-bit length.");
        byte[] framed = Join(new byte[] { 0, (byte)snappy.Length, (byte)(snappy.Length >> 8), (byte)(snappy.Length >> 16) }, snappy);
        using var output = new MemoryStream();
        using (var zip = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true)) {
            var entry = zip.CreateEntry("Index/Document.iwa", CompressionLevel.NoCompression);
            entry.LastWriteTime = new DateTimeOffset(2024, 1, 1, 0, 0, 0, TimeSpan.Zero);
            using Stream stream = entry.Open();
            stream.Write(framed);
        }
        return output.ToArray();
    }

    private static byte[] Pages(int units) => Join(
        Record(1, 10000, Reference(4, 2)),
        Record(2, 2001, String(3, string.Join("\n", Enumerable.Range(1, units).Select(Text)))));

    private static byte[] Keynote(int units) {
        var records = new List<byte[]> { Record(1, 1, Reference(2, 2)) };
        var nodes = new List<byte[]>();
        for (int i = 0; i < units; i++) {
            ulong node = (ulong)(10 + i * 4), slide = node + 1, shape = node + 2, storage = node + 3;
            nodes.Add(Reference(2, node));
            records.Add(Record(node, 4, Reference(2, slide)));
            records.Add(Record(slide, 5, Reference(5, shape)));
            records.Add(Record(shape, 2011, Join(Bytes(1, Bytes(1, Geometry())), Reference(2, storage))));
            records.Add(Record(storage, 2001, String(3, Text(i + 1))));
        }
        records.Add(Record(2, 2, Join(Bytes(3, Join(nodes.ToArray())),
            Bytes(4, Join(Float(1, 960), Float(2, 540))))));
        return Join(records.ToArray());
    }

    private static byte[] Geometry() => Bytes(1, Join(
        Bytes(1, Join(Float(1, 72), Float(2, 72))),
        Bytes(2, Join(Float(1, 240), Float(2, 60)))));

    private static byte[] Numbers(int units) {
        var records = new List<byte[]> {
            Record(1, 1, Reference(1, 2)),
            Record(2, 2, Join(String(1, "Sheet"), Reference(2, 10))),
            Record(10, 6000, Reference(2, 11))
        };
        var tiles = new List<byte[]>();
        for (int start = 0; start < units; start += 256) {
            ulong tile = (ulong)(100 + start / 256);
            var rows = new List<byte[]>();
            for (int row = 0; row < Math.Min(256, units - start); row++) {
                byte[] cell = new byte[20]; cell[0] = 5; cell[1] = 2;
                BinaryPrimitives.WriteUInt32LittleEndian(cell.AsSpan(8), 1u << 1);
                BinaryPrimitives.WriteInt64LittleEndian(cell.AsSpan(12), BitConverter.DoubleToInt64Bits(start + row + 1));
                rows.Add(Bytes(5, Join(Unsigned(1, (ulong)row), Bytes(6, cell), Bytes(7, new byte[] { 0, 0 }))));
            }
            records.Add(Record(tile, 6002, Join(rows.ToArray())));
            tiles.Add(Bytes(1, Join(Unsigned(1, (ulong)(start / 256)), Reference(2, tile))));
        }
        records.Add(Record(11, 6001, Join(Bytes(4, Bytes(3, Join(tiles.ToArray()))),
            Unsigned(6, (ulong)units), Unsigned(7, 1), String(8, "Values"))));
        return Join(records.ToArray());
    }

    private static byte[] Record(ulong id, ulong type, byte[] payload) {
        byte[] info = Join(Unsigned(1, id), Bytes(2, Join(Unsigned(1, type), Unsigned(3, (ulong)payload.Length))));
        return Join(Varint((ulong)info.Length), info, payload);
    }
    private static byte[] Reference(int field, ulong id) => Bytes(field, Unsigned(1, id));
    private static byte[] String(int field, string text) => Bytes(field, Encoding.UTF8.GetBytes(text));
    private static byte[] Unsigned(int field, ulong value) => Join(Varint((ulong)(field << 3)), Varint(value));
    private static byte[] Float(int field, float value) {
        byte[] data = new byte[4];
        BinaryPrimitives.WriteInt32LittleEndian(data, BitConverter.SingleToInt32Bits(value));
        return Join(Varint((ulong)((field << 3) | 5)), data);
    }
    private static byte[] Bytes(int field, byte[] data) => Join(Varint((ulong)((field << 3) | 2)), Varint((ulong)data.Length), data);
    private static byte[] LiteralHeader(int length) {
        int encoded = length - 1;
        if (length <= 60) return new[] { (byte)(encoded << 2) };
        int count = encoded <= 255 ? 1 : encoded <= 65535 ? 2 : encoded <= 0xffffff ? 3 : 4;
        var result = new byte[count + 1]; result[0] = (byte)((59 + count) << 2);
        for (int i = 0; i < count; i++) result[i + 1] = (byte)(encoded >> (8 * i));
        return result;
    }
    private static byte[] Varint(ulong value) {
        var result = new List<byte>();
        do { byte next = (byte)(value & 127); value >>= 7; result.Add(value == 0 ? next : (byte)(next | 128)); } while (value != 0);
        return result.ToArray();
    }
    private static byte[] Join(params byte[][] parts) {
        var result = new byte[checked(parts.Sum(part => part.Length))]; int offset = 0;
        foreach (byte[] part in parts) { part.CopyTo(result, offset); offset += part.Length; }
        return result;
    }
}
