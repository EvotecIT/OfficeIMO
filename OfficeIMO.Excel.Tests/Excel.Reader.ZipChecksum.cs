#if NET8_0_OR_GREATER
using System.IO.Compression;
using System.Text;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Reader_PreservesRuntimeZipChecksumValidation(bool corruptChecksum) {
        byte[] xml = Encoding.UTF8.GetBytes(
            "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
            + "<dimension ref=\"A1:A2\"/><sheetData><row r=\"1\"><c r=\"A1\" t=\"str\"><v>Id</v></c></row>"
            + "<row r=\"2\"><c r=\"A2\"><v>42</v></c></row>" + new string(' ', 1_100_000)
            + "</sheetData></worksheet>");
        using var packed = new MemoryStream();
        using (var compressor = new DeflateStream(packed, CompressionLevel.Optimal, leaveOpen: true)) {
            compressor.Write(xml);
        }
        byte[] package = CreateRawDeflateWorkbook(xml, packed.ToArray(), corruptChecksum);
        Exception? referenceError = Record.Exception(() => {
            using var archive = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
            using Stream entry = archive.GetEntry("xl/worksheets/sheet1.xml")!.Open();
            using var decoded = new MemoryStream();
            entry.CopyTo(decoded);
            Assert.Equal(xml, decoded.ToArray());
        });

        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.ZipChecksum.{Guid.NewGuid():N}.xlsx");
        try {
            File.WriteAllBytes(path, package);
            Exception? actual = Record.Exception(() => {
                using var reader = ExcelDocument.OpenReader(path);
                Assert.True(reader.Read());
                Assert.Equal(42, reader.GetInt32(0));
                Assert.False(reader.Read());
            });
            if (referenceError == null) Assert.Null(actual);
            else {
                Assert.IsType<InvalidDataException>(referenceError);
                Assert.IsType<InvalidDataException>(actual);
            }
        } finally {
            File.Delete(path);
        }
    }
}
#endif
