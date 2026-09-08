using System.IO.Compression;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    [Theory]
    [InlineData("utf-8")]
    [InlineData("utf-16")]
    [InlineData("utf-32")]
    public void SmallMetadataPreservesEncodingAndExtensionNames(string encodingName) {
        string path = CreateCompactFastPathWorkbook();
        try {
            const string entry = "xl/workbook.xml";
            XDocument workbook = XDocument.Parse(Encoding.UTF8.GetString(ReadZipEntry(path, entry)));
            XNamespace ns = workbook.Root!.Name.Namespace;
            workbook.Root.Element(ns + "sheets")!.Element(ns + "sheet")!.SetAttributeValue("name", "Łódź 🚀");
            XNamespace extension = "urn:officeimo:metadata-test";
            workbook.Root.Add(new XAttribute(XNamespace.Xmlns + "customPrefix", extension.NamespaceName));
            workbook.Root.Add(new XElement(extension + "customElement", new XAttribute("customAttribute", "漢")));
            using var output = new MemoryStream();
            using (var writer = XmlWriter.Create(output, new XmlWriterSettings { Encoding = Encoding.GetEncoding(encodingName) })) {
                workbook.Save(writer);
            }
            ReplaceZipEntry(path, entry, output.ToArray());

            Assert.Equal(new[] { "Łódź 🚀" }, XlsxTabularWorkbook.ReadSheetNames(path, new ExcelReadOptions()));
            Assert.Equal(new[] { "Łódź 🚀" }, ExcelDocument.GetSheetNames(path));
            AssertCompactNumericRows(path);
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void SmallMetadataStillRejectsDtdDeclarations() {
        string path = CreateCompactFastPathWorkbook();
        try {
            string xml = Encoding.UTF8.GetString(ReadZipEntry(path, "xl/workbook.xml"));
            int rootStart = xml.IndexOf("<workbook", StringComparison.Ordinal);
            xml = xml.Insert(rootStart, "<!DOCTYPE workbook [<!ENTITY name 'unexpected'>]>");
            ReplaceZipEntry(path, "xl/workbook.xml", Encoding.UTF8.GetBytes(xml));
            Assert.Throws<XlsxTabularFastPathNotSupportedException>(() =>
                XlsxTabularWorkbook.ReadSheetNames(path, new ExcelReadOptions()));
        } finally {
            File.Delete(path);
        }
    }
}

public class OpenXmlSmallPartTests {
    [Theory]
    [InlineData(0)]
    [InlineData(16)]
    [InlineData(4095)]
    [InlineData(4096)]
    [InlineData(4097)]
    public void PartReadsExcludeUnusedBufferCapacity(int length) {
        byte[] content = Enumerable.Range(0, length).Select(index => (byte)(index % 251)).ToArray();
        using var parts = OpenXmlPackagePartBufferReader.TryOpen(Package("xl/test.xml", content))!;
        using Stream part = parts.OpenPart("xl/test.xml", 4097);
        using var output = new MemoryStream();
        part.CopyTo(output);
        Assert.Equal(content, output.ToArray());
        Assert.Equal(-1, part.ReadByte());
    }

    [Theory]
    [InlineData(140)]
    [InlineData(4096)]
    public void BufferedPartsRejectTruncatedInput(int declaredLength) {
        byte[] package = Package("xl/test.xml", Enumerable.Repeat((byte)'x', 128).ToArray());
        int directory = -1;
        for (int offset = 0; offset <= package.Length - 46; offset++) {
            if (BitConverter.ToUInt32(package, offset) == 0x02014B50U) { directory = offset; break; }
        }
        Assert.True(directory >= 0);
        Array.Copy(BitConverter.GetBytes(declaredLength), 0, package, directory + 24, 4);
        using var parts = OpenXmlPackagePartBufferReader.TryOpen(package)!;
        Exception? error = Record.Exception(() => { using Stream part = parts.OpenPart("xl/test.xml", 4096); });
        Assert.True(error is InvalidDataException or EndOfStreamException, error?.ToString() ?? "The corrupt part was accepted.");
    }

    [Theory]
    [InlineData("xl/sheet-A_1.~.xml", "xl/sheet-A_1.~.xml")]
    [InlineData("xl/.../sheet.xml", "xl/.../sheet.xml")]
    [InlineData("xl/sheet one.xml", "xl/sheet%20one.xml")]
    [InlineData("xl/sheet%20one.xml", "xl/sheet one.xml")]
    [InlineData("xl/Żółć.xml", "xl/%C5%BB%C3%B3%C5%82%C4%87.xml")]
    [InlineData("xl/sheet%2520one.xml", "xl/sheet%2520one.xml")]
    public void CanonicalAndEncodedPartNamesResolveToTheirBytes(string storedName, string requestedName) {
        byte[] expected = Encoding.UTF8.GetBytes(storedName);
        using var parts = OpenXmlPackagePartBufferReader.TryOpen(Package(storedName, expected))!;
        using Stream part = parts.OpenPart(requestedName, 4096);
        using var output = new MemoryStream();
        part.CopyTo(output);
        Assert.Equal(expected, output.ToArray());
    }

    [Theory]
    [InlineData("xl/../sheet.xml")]
    [InlineData("xl/%2E%2E/sheet.xml")]
    [InlineData("xl//sheet.xml")]
    [InlineData("xl/bad%2.xml")]
    public void InvalidPartNamesAreStillRejected(string name) {
        Assert.Throws<OpenXmlPackagePartIndexException>(() =>
            OpenXmlPackagePartBufferReader.TryOpen(Package(name, new byte[] { 42 })));
    }

    [Fact]
    public void DisposedPooledPartCannotExposeReturnedBytes() {
        byte[] bytes = System.Buffers.ArrayPool<byte>.Shared.Rent(16);
        using var stream = new OpenXmlPooledPartStream(bytes, 1);
        stream.Dispose();
        Assert.Throws<ObjectDisposedException>(() => stream.ToArray());
        Assert.Throws<ObjectDisposedException>(() => stream.ReadByte());
    }

    private static byte[] Package(string name, byte[] content) {
        using var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true)) {
            using Stream entry = archive.CreateEntry(name, CompressionLevel.Fastest).Open();
            entry.Write(content, 0, content.Length);
        }
        return stream.ToArray();
    }
}

public class OpenXmlReadNameTableTests {
    [Fact]
    public void UnknownNamesStayLocalWhileSchemaNamesCanBeReadConcurrently() {
        Parallel.For(0, 100, index => {
            var first = new OpenXmlReadNameTable();
            var second = new OpenXmlReadNameTable();
            string name = "custom-" + index;
            string atom = first.Add(name.ToCharArray(), 0, name.Length);
            Assert.Same(atom, first.Add(new string(name.ToCharArray())));
            Assert.Null(second.Get(name));
            Assert.Same(first.Add("workbook"), second.Add("workbook".ToCharArray(), 0, 8));
        });
    }
}
