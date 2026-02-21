using OfficeIMO.Reader;
using OfficeIMO.Reader.Xml;
using System.Text;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderXmlModularTests {
    [Fact]
    public void DocumentReaderXml_ReadXml_ParsesXmlIntoStructuredChunks() {
        var path = Path.Combine(Path.GetTempPath(), "officeimo-xml-" + Guid.NewGuid().ToString("N") + ".xml");
        try {
            File.WriteAllText(path,
                "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
                "<catalog><book id=\"b1\"><title>One</title></book><book id=\"b2\"><title>Two</title></book></catalog>");

            var chunks = DocumentReaderXmlExtensions.ReadXml(
                path,
                xmlOptions: new XmlReadOptions {
                    ChunkRows = 2,
                    IncludeMarkdown = true
                }).ToList();

            Assert.NotEmpty(chunks);
            Assert.All(chunks, c => Assert.Equal(ReaderInputKind.Xml, c.Kind));
            Assert.Contains(chunks, c => (c.Text ?? string.Empty).Contains("catalog[1]/book[1]/@id", StringComparison.Ordinal));
            Assert.Contains(chunks, c => c.Tables != null && c.Tables.Count > 0 && c.Tables[0].Columns.Contains("Type", StringComparer.Ordinal));
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void DocumentReaderXml_ReadXmlStream_NonSeekable_EnforcesMaxInputBytes() {
        var xml =
            "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
            "<catalog><book id=\"b1\"><title>One</title></book><book id=\"b2\"><title>Two</title></book></catalog>";
        using var stream = new NonSeekableReadStream(Encoding.UTF8.GetBytes(xml));

        var ex = Assert.Throws<IOException>(() => DocumentReaderXmlExtensions.ReadXml(
            stream,
            sourceName: "books.xml",
            readerOptions: new ReaderOptions { MaxInputBytes = 16 },
            xmlOptions: new XmlReadOptions {
                ChunkRows = 2,
                IncludeMarkdown = true
            }).ToList());

        Assert.Contains("Input exceeds MaxInputBytes", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void DocumentReaderXml_ReadXml_EmitsWarningForMalformedXml() {
        var path = Path.Combine(Path.GetTempPath(), "officeimo-bad-xml-" + Guid.NewGuid().ToString("N") + ".xml");
        try {
            File.WriteAllText(path, "<root><broken></root>");

            var chunks = DocumentReaderXmlExtensions.ReadXml(path).ToList();

            Assert.NotEmpty(chunks);
            Assert.Contains(chunks, c =>
                c.Kind == ReaderInputKind.Xml &&
                (c.Warnings?.Any(w => w.Contains("XML parse error", StringComparison.OrdinalIgnoreCase)) ?? false));
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }
}
