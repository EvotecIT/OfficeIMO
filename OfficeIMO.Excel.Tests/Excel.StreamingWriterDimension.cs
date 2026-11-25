using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    /// <summary>
    /// A streamed table declares its dimension, filled in after the rows when the destination can be rewritten, so
    /// readers know its columns before the first row. The package stays valid and its CRC covers the final bytes.
    /// </summary>
    [Theory]
    [InlineData(null, "A1", 10, "A1:C11")]
    [InlineData("Report", "C4", 10, "C4:E15")]
    [InlineData(null, "A1", 0, "A1:C1")]
    public void StreamingWriter_DeclaresTheWrittenRangeOnRewritableOutput(string? title, string address, int rows, string expected) {
        using var output = new MemoryStream();
        var options = new ExcelWriterOptions { Title = title, Address = address };
        using (var writer = ExcelDocument.CreateWriter(output, options, leaveOpen: true)) {
            writer.WriteTable("Data", Enumerable.Range(1, rows).Select(row => (Id: row, Name: "Item " + row, Amount: row * 1.5)),
                new[] { "Id", "Name", "Amount" }, static (row, item) => row.Write(item.Id).Write(item.Name).Write(item.Amount));
        }
        byte[] package = output.ToArray();

        Assert.Equal(expected, ReadStreamedDimension(package));
        AssertStreamedPackageIsValid(package);
    }

    /// <summary>A forward-only destination cannot be rewritten, so the worksheet has no dimension and stays valid.</summary>
    [Fact]
    public void StreamingWriter_OmitsTheDimensionOnForwardOnlyOutput() {
        var output = new ForwardOnlyPackageOutput();
        using (var writer = ExcelDocument.CreateWriter(output, new ExcelWriterOptions(), leaveOpen: true)) {
            writer.WriteTable("Data", Enumerable.Range(1, 10), new[] { "Id" }, static (row, item) => row.Write(item));
        }
        byte[] package = output.ToArray();

        Assert.Null(ReadStreamedDimension(package));
        AssertStreamedPackageIsValid(package);
    }

    private static string? ReadStreamedDimension(byte[] package) {
        using var archive = new ZipArchive(new MemoryStream(package, writable: false), ZipArchiveMode.Read);
        using Stream sheet = archive.GetEntry("xl/worksheets/sheet1.xml")!.Open();
        XNamespace main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        return XDocument.Load(sheet).Root!.Element(main + "dimension")?.Attribute("ref")?.Value;
    }

    private static void AssertStreamedPackageIsValid(byte[] package) {
        Dictionary<string, uint> localCrcs = ReadLocalHeaderCrcs(package);
        using (var archive = new ZipArchive(new MemoryStream(package, writable: false), ZipArchiveMode.Read)) {
            foreach (ZipArchiveEntry entry in archive.Entries) {
                using var content = new MemoryStream();
                using (Stream input = entry.Open()) input.CopyTo(content);
                uint crc = StreamedEntryCrc32(content.ToArray());
                Assert.Equal(entry.Crc32, crc);
                // Readers that trust the local header must see the same value as the central directory.
                if (localCrcs.TryGetValue(entry.FullName, out uint localCrc)) Assert.Equal(crc, localCrc);
            }
        }
        using var spreadsheet = SpreadsheetDocument.Open(new MemoryStream(package, writable: false), false);
        Assert.Empty(new OpenXmlValidator().Validate(spreadsheet));
    }

    /// <summary>The CRC-32 in each local header that carries one (no data descriptor).</summary>
    private static Dictionary<string, uint> ReadLocalHeaderCrcs(byte[] package) {
        var crcs = new Dictionary<string, uint>(StringComparer.Ordinal);
        int offset = 0;
        while (offset + 30 <= package.Length && BitConverter.ToUInt32(package, offset) == 0x04034B50U) {
            ushort flags = BitConverter.ToUInt16(package, offset + 6);
            uint compressedSize = BitConverter.ToUInt32(package, offset + 18);
            int nameLength = BitConverter.ToUInt16(package, offset + 26);
            int extraLength = BitConverter.ToUInt16(package, offset + 28);
            string name = Encoding.UTF8.GetString(package, offset + 30, nameLength);
            if ((flags & 8) != 0) break;
            crcs[name] = BitConverter.ToUInt32(package, offset + 14);
            offset += 30 + nameLength + extraLength + checked((int)compressedSize);
        }
        return crcs;
    }

    private static uint StreamedEntryCrc32(byte[] data) {
        uint crc = 0xFFFFFFFFU;
        foreach (byte value in data) {
            crc ^= value;
            for (int bit = 0; bit < 8; bit++) crc = (crc >> 1) ^ (0xEDB88320U & (uint)-(int)(crc & 1));
        }
        return crc ^ 0xFFFFFFFFU;
    }

    private sealed class ForwardOnlyPackageOutput : Stream {
        private readonly MemoryStream _inner = new();
        public byte[] ToArray() => _inner.ToArray();
        public override bool CanRead => false;
        public override bool CanSeek => false;
        public override bool CanWrite => true;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override void Flush() { }
        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => _inner.Write(buffer, offset, count);
    }
}
