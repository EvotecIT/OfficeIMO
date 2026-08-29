using System.Text;

namespace OfficeIMO.LegacyImport.Tests;

internal static class LegacyFixtureFactory {
    internal static byte[] WordStar() => Encoding.ASCII.GetBytes("First paragraph\r\n- List item\r\n")
        .Select((value, index) => index % 3 == 0 && value >= 32 ? (byte)(value | 0x80) : value).ToArray();

    internal static byte[] WordPerfect() {
        byte[] text = Encoding.ASCII.GetBytes("Recovered WordPerfect text\r\nSecond paragraph");
        byte[] data = new byte[16 + text.Length];
        data[0] = 0xFF; data[1] = 0x57; data[2] = 0x50; data[3] = 0x43;
        data[4] = 16;
        Buffer.BlockCopy(text, 0, data, 16, text.Length);
        return data;
    }

    internal static byte[] AmiPro() => Encoding.ASCII.GetBytes("[ver]\n[sty]\nAmi Pro paragraph\n- Ami list");

    internal static byte[] WordPro() => CompoundLike("Lotus Word Pro recovered text");

    internal static byte[] WorksWord() {
        byte[] data = new byte[96];
        data[0] = 2; data[1] = 0xFE;
        Encoding.ASCII.GetBytes("Works recovered paragraph").CopyTo(data, 16);
        return data;
    }

    internal static byte[] Write(bool write) {
        byte[] data = new byte[220];
        data[0] = 0x31; data[1] = 0xBE; data[5] = 0xAB; data[96] = write ? (byte)1 : (byte)0;
        Encoding.ASCII.GetBytes(write ? "Write recovered paragraph" : "Word DOS recovered paragraph").CopyTo(data, 128);
        return data;
    }

    internal static byte[] Wk(byte product0 = 0x04, byte product1 = 0x04, bool includeFormulaAndChart = true, byte cellFormat = 0) {
        using var stream = new MemoryStream();
        using var writer = new BinaryWriter(stream, Encoding.ASCII, leaveOpen: true);
        Record(writer, 0x0000, new[] { product0, product1 });
        Record(writer, 0x000F, LabelPayload(0, 0, (byte)'\'', Encoding.ASCII.GetBytes("Name\0"), cellFormat));
        Record(writer, 0x000D, CellPayload(1, 0, BitConverter.GetBytes((short)42), cellFormat));
        if (includeFormulaAndChart) {
            Record(writer, 0x0010, CellPayload(2, 0, BitConverter.GetBytes(84d).Concat(new byte[] { 0x01 }).ToArray(), cellFormat));
            Record(writer, 0x002D, new byte[] { 1, 2, 3 });
        }
        Record(writer, 0x0001, Array.Empty<byte>());
        writer.Flush();
        return stream.ToArray();
    }

    internal static byte[] Multiplan() {
        byte[] body = Encoding.ASCII.GetBytes("Name\tValue\nA\t12\n");
        return new byte[] { 0x08, 0xE7 }.Concat(body).ToArray();
    }

    internal static byte[] CompoundSheet() => CompoundLike("Name\tValue\nA\t10");

    private static byte[] CompoundLike(string text) {
        byte[] data = new byte[128];
        new byte[] { 0xD0, 0xCF, 0x11, 0xE0 }.CopyTo(data, 0);
        Encoding.ASCII.GetBytes(text).CopyTo(data, 32);
        return data;
    }

    private static byte[] CellPayload(ushort column, ushort row, byte[] value, byte format = 0) =>
        new[] { format, (byte)column, (byte)(column >> 8), (byte)row, (byte)(row >> 8) }.Concat(value).ToArray();

    private static byte[] LabelPayload(ushort column, ushort row, byte marker, byte[] value, byte format = 0) =>
        CellPayload(column, row, new[] { marker }.Concat(value).ToArray(), format);

    private static void Record(BinaryWriter writer, ushort type, byte[] payload) {
        writer.Write(type);
        writer.Write((ushort)payload.Length);
        writer.Write(payload);
    }
}
