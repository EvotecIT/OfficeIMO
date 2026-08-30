using System.Text;

namespace OfficeIMO.LegacyImport.Tests;

internal static class LegacyFixtureFactory {
    internal static byte[] WordStar() {
        var bytes = new List<byte>();
        bytes.AddRange(Encoding.ASCII.GetBytes("First "));
        bytes.Add(0x02);
        bytes.AddRange(Encoding.ASCII.GetBytes("paragraph"));
        bytes.Add(0x02);
        bytes.AddRange(new byte[] { 0x0D, 0x0A });
        bytes.AddRange(Encoding.ASCII.GetBytes(".PA\r\n- List item\r\n"));
        bytes.AddRange(WordStarSequence(0x06, "Recovered comment"));
        bytes.Add(0x1A);
        for (int index = 0; index < bytes.Count; index++) if (bytes[index] >= 32 && bytes[index] <= 126 && index % 3 == 0) bytes[index] |= 0x80;
        return bytes.ToArray();
    }

    internal static byte[] WordStarWithGraphics() {
        var bytes = new List<byte>(Encoding.ASCII.GetBytes("Illustrated\r\n"));
        bytes.AddRange(WordStarSequence(0x10, "C:\\IMAGES\\FIGURE.PCX"));
        bytes.Add(0x1A);
        return bytes.ToArray();
    }

    internal static byte[] WordPerfect() {
        byte[] text = Encoding.ASCII.GetBytes("Recovered WordPerfect text\r\nSecond paragraph");
        byte[] data = new byte[16 + text.Length];
        data[0] = 0xFF; data[1] = 0x57; data[2] = 0x50; data[3] = 0x43;
        data[4] = 16;
        Buffer.BlockCopy(text, 0, data, 16, text.Length);
        return data;
    }

    internal static byte[] AmiPro() => Encoding.ASCII.GetBytes(
        "[ver]\n4\n[tag]\nBody Text\n0\n[fnt]\nArial\n240\n255\n16385\n[algn]\n1\n0\n0\n0\n0\n[spc]\n1\n240\n0\n0\n0\n[brk]\n16\n[edoc]\n@Body Text@Ami Pro <+!>bold<-!> paragraph\n\n<+B>- Ami list\n");

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

    internal static byte[] Wk(byte product0 = 0x04, byte product1 = 0x04, bool includeFormulaAndChart = true, byte cellFormat = 0, byte[]? formulaTokens = null, ushort? declaredFormulaLength = null, bool includeBlank = false, ushort? extraRecordType = null) {
        using var stream = new MemoryStream();
        using var writer = new BinaryWriter(stream, Encoding.ASCII, leaveOpen: true);
        Record(writer, 0x0000, new[] { product0, product1 });
        Record(writer, 0x000B, NamePayload("Input", 0, 0, 1, 0));
        Record(writer, 0x000F, LabelPayload(0, 0, (byte)'\'', Encoding.ASCII.GetBytes("Name\0"), cellFormat));
        Record(writer, 0x000D, CellPayload(1, 0, BitConverter.GetBytes((short)42), cellFormat));
        if (includeBlank) Record(writer, 0x000C, CellPayload(3, 0, Array.Empty<byte>(), cellFormat));
        if (includeFormulaAndChart) {
            byte[] tokens = formulaTokens ?? new byte[] { 0x01, 0, 0, 0, 0, 0x01, 1, 0, 0, 0, 0x09, 0x03 };
            Record(writer, 0x0010, CellPayload(2, 0, BitConverter.GetBytes(84d).Concat(BitConverter.GetBytes(declaredFormulaLength ?? (ushort)tokens.Length)).Concat(tokens).ToArray(), cellFormat));
            Record(writer, 0x002D, new byte[] { 1, 2, 3 });
        }
        if (extraRecordType.HasValue) Record(writer, extraRecordType.Value, new byte[] { 1, 2, 3 });
        Record(writer, 0x0001, Array.Empty<byte>());
        writer.Flush();
        return stream.ToArray();
    }

    internal static byte[] WkWithNames(params string[] names) {
        using var stream = new MemoryStream();
        using var writer = new BinaryWriter(stream, Encoding.ASCII, leaveOpen: true);
        Record(writer, 0x0000, new byte[] { 0x04, 0x04 });
        foreach (string name in names) Record(writer, 0x000B, NamePayload(name, 0, 0, 0, 0));
        Record(writer, 0x000D, CellPayload(0, 0, BitConverter.GetBytes((short)1)));
        Record(writer, 0x0001, Array.Empty<byte>());
        writer.Flush();
        return stream.ToArray();
    }

    internal static byte[] Wq2WithTruncatedFormulaEnvelope() {
        using var stream = new MemoryStream();
        using var writer = new BinaryWriter(stream, Encoding.ASCII, leaveOpen: true);
        Record(writer, 0x0000, new byte[] { 0x21, 0x51 });
        byte[] payload = new byte[14];
        BitConverter.GetBytes(12d).CopyTo(payload, 6);
        Record(writer, 0x0010, payload);
        Record(writer, 0x0001, Array.Empty<byte>());
        writer.Flush();
        return stream.ToArray();
    }

    internal static byte[] WkMultiSheet() {
        using var stream = new MemoryStream();
        using var writer = new BinaryWriter(stream, Encoding.ASCII, leaveOpen: true);
        Record(writer, 0x0000, new byte[] { 0x04, 0x04 });
        Record(writer, 0x000D, CellPayload(0, 0, BitConverter.GetBytes((short)1)));
        Record(writer, 0x000D, CellPayload(0, 0, BitConverter.GetBytes((short)2), sheet: 1));
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

    private static byte[] CellPayload(byte column, ushort row, byte[] value, byte format = 0, byte sheet = 0) =>
        new[] { format, column, sheet, (byte)row, (byte)(row >> 8) }.Concat(value).ToArray();

    private static byte[] LabelPayload(byte column, ushort row, byte marker, byte[] value, byte format = 0) =>
        CellPayload(column, row, new[] { marker }.Concat(value).ToArray(), format);

    private static byte[] NamePayload(string name, ushort firstColumn, ushort firstRow, ushort lastColumn, ushort lastRow) {
        byte[] payload = new byte[24];
        byte[] encoded = Encoding.ASCII.GetBytes(name);
        payload[0] = (byte)Math.Min(15, encoded.Length);
        encoded.AsSpan(0, payload[0]).CopyTo(payload.AsSpan(1));
        BitConverter.GetBytes(firstColumn).CopyTo(payload, 16);
        BitConverter.GetBytes(firstRow).CopyTo(payload, 18);
        BitConverter.GetBytes(lastColumn).CopyTo(payload, 20);
        BitConverter.GetBytes(lastRow).CopyTo(payload, 22);
        return payload;
    }

    private static byte[] WordStarSequence(byte type, string text) {
        byte[] payload = Encoding.ASCII.GetBytes(text);
        int totalLength = payload.Length + 7;
        ushort count = (ushort)(totalLength - 3);
        return new[] { (byte)0x1D, (byte)count, (byte)(count >> 8), type }
            .Concat(payload)
            .Concat(new[] { (byte)count, (byte)(count >> 8), (byte)0x1D })
            .ToArray();
    }

    private static void Record(BinaryWriter writer, ushort type, byte[] payload) {
        writer.Write(type);
        writer.Write((ushort)payload.Length);
        writer.Write(payload);
    }
}
