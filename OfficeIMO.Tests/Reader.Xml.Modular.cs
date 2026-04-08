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
            Assert.All(chunks, c => {
                Assert.False(string.IsNullOrWhiteSpace(c.SourceId));
                Assert.False(string.IsNullOrWhiteSpace(c.SourceHash));
                Assert.False(string.IsNullOrWhiteSpace(c.ChunkHash));
                Assert.True(c.TokenEstimate.HasValue && c.TokenEstimate.Value >= 1);
                Assert.True(c.SourceLengthBytes.HasValue && c.SourceLengthBytes.Value > 0);
                Assert.True(c.SourceLastWriteUtc.HasValue);
            });
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

    [Fact]
    public void DocumentReaderXml_ReadXml_PreservesCDataAsElementText() {
        var path = Path.Combine(Path.GetTempPath(), "officeimo-cdata-xml-" + Guid.NewGuid().ToString("N") + ".xml");
        try {
            File.WriteAllText(path,
                "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
                "<catalog><book><summary><![CDATA[Line <b>One</b> & more]]></summary></book></catalog>");

            var chunk = Assert.Single(DocumentReaderXmlExtensions.ReadXml(
                path,
                xmlOptions: new XmlReadOptions {
                    ChunkRows = 10,
                    IncludeMarkdown = true
                }));

            Assert.Contains("catalog[1]/book[1]/summary[1] | element | Line <b>One</b> & more", chunk.Text ?? string.Empty, StringComparison.Ordinal);
            Assert.Contains("Line <b>One</b> & more", chunk.Markdown ?? string.Empty, StringComparison.Ordinal);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void DocumentReaderXml_ReadXml_PreservesQualifiedNamesInPaths() {
        var path = Path.Combine(Path.GetTempPath(), "officeimo-ns-xml-" + Guid.NewGuid().ToString("N") + ".xml");
        try {
            File.WriteAllText(path,
                "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
                "<root xmlns:a=\"urn:alpha\" xmlns:b=\"urn:beta\">" +
                "<a:item code=\"A1\">One</a:item>" +
                "<b:item b:code=\"B1\">Two</b:item>" +
                "</root>");

            var chunk = Assert.Single(DocumentReaderXmlExtensions.ReadXml(
                path,
                xmlOptions: new XmlReadOptions {
                    ChunkRows = 20,
                    IncludeMarkdown = true
                }));

            Assert.Contains("root[1]/a:item[1] | element | One", chunk.Text ?? string.Empty, StringComparison.Ordinal);
            Assert.Contains("root[1]/b:item[1] | element | Two", chunk.Text ?? string.Empty, StringComparison.Ordinal);
            Assert.Contains("root[1]/b:item[1]/@b:code | attribute | B1", chunk.Text ?? string.Empty, StringComparison.Ordinal);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }
}
