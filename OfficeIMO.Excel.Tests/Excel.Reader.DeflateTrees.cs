#if NET8_0_OR_GREATER
using System.Buffers.Binary;
using System.IO.Compression;
using System.Text;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    [Theory]
    [InlineData("complete-distance", true)]
    [InlineData("empty-distance", true)]
    [InlineData("end-only", true)]
    [InlineData("maximum-distance-count", true)]
    [InlineData("distance", false)]
    [InlineData("literal", false)]
    [InlineData("code-length", false)]
    [InlineData("distance-count", false)]
    public void Reader_ValidatesLaterDynamicHuffmanTrees(string treeKind, bool valid) {
        const string closing = "</sheetData></worksheet>";
        string opening = "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
            + "<dimension ref=\"A1:A2\"/><sheetData><row r=\"1\"><c r=\"A1\" t=\"str\"><v>Id</v></c></row>"
            + "<row r=\"2\"><c r=\"A2\"><v>42</v></c></row>";
        bool endOnly = treeKind == "end-only";
        byte[] prefix = Encoding.UTF8.GetBytes((opening + (endOnly ? closing : "")).PadRight(ushort.MaxValue));
        byte[] tail = Convert.FromHexString(DeflateTreeTail(treeKind));
        using var packed = new MemoryStream();
        using (var writer = new BinaryWriter(packed, Encoding.UTF8, leaveOpen: true)) {
            writer.Write((byte)0); // A non-final stored block keeps BCL's first 256 bytes independent of the later tree.
            writer.Write(ushort.MaxValue);
            writer.Write((ushort)0);
            writer.Write(prefix);
        }
        var bits = new DeflateTreeFixtureBits(packed);
        bits.WriteBits(0, 1);
        bits.WriteBits(1, 2); // Fixed Huffman block: repeat the established space history.
        for (int index = 0; index < 4096; index++) {
            bits.WriteCode(197, 8); // Length 258, followed by distance 1.
            bits.WriteBits(0, 5);
        }
        bits.WriteBits(0, 7); // End of fixed block.
        bits.WriteBits(0, 3); // An empty stored block aligns the following dynamic header.
        bits.Align();
        packed.Write(new byte[] { 0, 0, 0xff, 0xff });
        packed.Write(tail);
        byte[] compressed = packed.ToArray();
        byte[] expected = Encoding.UTF8.GetBytes(Encoding.UTF8.GetString(prefix)
            + new string(' ', 4096 * 258) + (endOnly ? "" : closing));

        Exception? referenceError = Record.Exception(() => {
            using var input = new DeflateStream(new MemoryStream(compressed), CompressionMode.Decompress);
            using var decoded = new MemoryStream();
            input.CopyTo(decoded);
            Assert.Equal(expected, decoded.ToArray());
        });
        if (valid) Assert.Null(referenceError);
        else Assert.IsType<InvalidDataException>(referenceError);

        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.DeflateTree.{Guid.NewGuid():N}.xlsx");
        try {
            File.WriteAllBytes(path, DeflateTreeWorkbook(expected, compressed));
            Exception? actual = Record.Exception(() => {
                using var reader = ExcelDocument.OpenReader(path);
                Assert.True(reader.Read());
                Assert.Equal(42, reader.GetInt32(0));
                Assert.False(reader.Read());
            });
            if (valid) Assert.Null(actual);
            else Assert.IsType<InvalidDataException>(actual);
        } finally {
            File.Delete(path);
        }
    }

    // Independent raw dynamic blocks emit only the closing XML, or just end-of-block.
    // The three malformed cases leave unused codes in the named alphabet.
    private static string DeflateTreeTail(string kind) => kind switch {
        "complete-distance" => "05c001401000000090aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaeaf1d03b5b989aba8818ba18f2f1d0bbdb3b593b5b989abaf0fd07",
        "empty-distance" => "0560004450555555555555555555555555555555555555555555555555555555555555555555555555555555555555555555555555555555555555555555555555555555f578e89d2d4c4d5d440c5d0cf978e8dded9dac9d2d4c4d5df8fe03",
        "end-only" => "05c001040000000010000000000000000000000000000000000000000000000000000000000000008000",
        "distance" => "0580014010000040aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaac743ef6c616aea2262e862c8c743ef6eef64ed6c616aeac2f71f",
        "literal" => "05c001000800000090fffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff0a0e7cc8229532e44187261880f0f7aeeec39b1e6cc8229532ef80200",
        "code-length" => "05c001401000000020555555555555555555555555555555555555555555555555555555555555555555555555555555555555555555555555555555555555555555555555555555b5e0a177b630357511317431e4e3a177b777b276b6303575e1fb0f",
        "maximum-distance-count" => "05dd014810000000a0aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaea050000000000008087ded9c2d4d445c4d0c5908f87dedddec9dad9c2d4d485ef3f",
        "distance-count" => "05de014810000000a0aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaea05000000000000001e7a670b535317114317433e1e7a777b276b670b535317beff",
        _ => throw new ArgumentOutOfRangeException(nameof(kind))
    };

    private static byte[] DeflateTreeWorkbook(byte[] xml, byte[] compressed) {
        const string main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        const string relationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        const string package = "http://schemas.openxmlformats.org/package/2006/relationships";
        const string worksheetName = "xl/worksheets/sheet1.xml";
        using var bytes = new MemoryStream();
        using (var zip = new ZipArchive(bytes, ZipArchiveMode.Create, leaveOpen: true)) {
            Write("[Content_Types].xml", "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
                + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
                + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
                + "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>"
                + "<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/></Types>");
            Write("_rels/.rels", $"<Relationships xmlns=\"{package}\"><Relationship Id=\"rId1\" Type=\"{relationships}/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>");
            Write("xl/workbook.xml", $"<workbook xmlns=\"{main}\" xmlns:r=\"{relationships}\"><sheets><sheet name=\"Data\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>");
            Write("xl/_rels/workbook.xml.rels", $"<Relationships xmlns=\"{package}\"><Relationship Id=\"rId1\" Type=\"{relationships}/worksheet\" Target=\"worksheets/sheet1.xml\"/></Relationships>");
            using Stream sheet = zip.CreateEntry(worksheetName, CompressionLevel.NoCompression).Open();
            sheet.Write(xml);

            void Write(string name, string text) {
                using Stream entry = zip.CreateEntry(name, CompressionLevel.NoCompression).Open();
                entry.Write(Encoding.UTF8.GetBytes(text));
            }
        }
        byte[] original = bytes.ToArray();
        int directory = checked((int)BinaryPrimitives.ReadUInt32LittleEndian(original.AsSpan(original.Length - 6)));
        int central = directory;
        while (true) {
            int nameLength = BinaryPrimitives.ReadUInt16LittleEndian(original.AsSpan(central + 28));
            if (Encoding.UTF8.GetString(original, central + 46, nameLength) == worksheetName) break;
            central += 46 + nameLength + BinaryPrimitives.ReadUInt16LittleEndian(original.AsSpan(central + 30))
                + BinaryPrimitives.ReadUInt16LittleEndian(original.AsSpan(central + 32));
        }
        int local = checked((int)BinaryPrimitives.ReadUInt32LittleEndian(original.AsSpan(central + 42)));
        int start = local + 30 + BinaryPrimitives.ReadUInt16LittleEndian(original.AsSpan(local + 26))
            + BinaryPrimitives.ReadUInt16LittleEndian(original.AsSpan(local + 28));
        // The worksheet is last. Replace its payload and optional descriptor with a fully specified local header.
        byte[] result = original.Take(start).Concat(compressed).Concat(original.Skip(directory)).ToArray();
        int delta = compressed.Length - (directory - start);
        ushort flags = (ushort)(BinaryPrimitives.ReadUInt16LittleEndian(original.AsSpan(central + 8)) & ~8);
        BinaryPrimitives.WriteUInt16LittleEndian(result.AsSpan(local + 4), 20);
        BinaryPrimitives.WriteUInt16LittleEndian(result.AsSpan(local + 6), flags);
        BinaryPrimitives.WriteUInt16LittleEndian(result.AsSpan(local + 8), 8);
        original.AsSpan(central + 16, 4).CopyTo(result.AsSpan(local + 14));
        original.AsSpan(central + 24, 4).CopyTo(result.AsSpan(local + 22));
        BinaryPrimitives.WriteUInt32LittleEndian(result.AsSpan(local + 18), (uint)compressed.Length);
        BinaryPrimitives.WriteUInt16LittleEndian(result.AsSpan(central + delta + 6), 20);
        BinaryPrimitives.WriteUInt16LittleEndian(result.AsSpan(central + delta + 8), flags);
        BinaryPrimitives.WriteUInt16LittleEndian(result.AsSpan(central + delta + 10), 8);
        BinaryPrimitives.WriteUInt32LittleEndian(result.AsSpan(central + delta + 20), (uint)compressed.Length);
        BinaryPrimitives.WriteUInt32LittleEndian(result.AsSpan(result.Length - 6), (uint)(directory + delta));
        return result;
    }

    private sealed class DeflateTreeFixtureBits {
        private readonly Stream _output;
        private int _value, _count;
        internal DeflateTreeFixtureBits(Stream output) => _output = output;
        internal void WriteBits(int value, int count) {
            for (int bit = 0; bit < count; bit++) {
                _value |= ((value >> bit) & 1) << _count++;
                if (_count == 8) Align();
            }
        }
        internal void WriteCode(int value, int count) {
            for (int bit = count - 1; bit >= 0; bit--) WriteBits((value >> bit) & 1, 1);
        }
        internal void Align() {
            if (_count == 0) return;
            _output.WriteByte((byte)_value);
            _value = _count = 0;
        }
    }
}
#endif
