using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordXmlInterchangeTests {
    [Fact]
    public async Task FlatOpcLoadAndExportRetainEditsWithoutAssociatingWritableSource() {
        byte[] original = Encoding.UTF8.GetBytes(FlatOpc());
        using var source = new MemoryStream();
        source.Write(original, 0, original.Length);
        source.Position = 0;

        using WordDocument document = await WordDocument.LoadAsync(source);
        Assert.Equal(WordFileFormat.FlatOpc, document.SourceFormat);
        Assert.Equal("Word.FlatOpc", document.SourceFormatDescriptor.Id);
        Assert.Single(document.Paragraphs).SetText("Edited package");
        Assert.Throws<InvalidOperationException>(() => document.Save());
        Assert.Throws<InvalidOperationException>(() => document.Save(source));
        await Assert.ThrowsAsync<InvalidOperationException>(() => document.SaveAsync(source));
        Assert.Equal(original, source.ToArray());

        byte[] exported = document.ExportFlatOpc();
        using var exportedStream = new MemoryStream(exported);
        XDocument xml = XDocument.Load(exportedStream);
        XNamespace pkg = "http://schemas.microsoft.com/office/2006/xmlPackage";
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        Assert.Equal(pkg + "package", xml.Root.Name);
        Assert.Equal("Edited package", Assert.Single(xml.Descendants(w + "t")).Value);
        Assert.Contains(xml.Root.Elements(pkg + "part"), part => (string?)part.Attribute(pkg + "name") == "/_rels/.rels");
        exportedStream.Position = 0;
        using WordDocument reopened = WordDocument.Load(exportedStream);
        Assert.Equal("Edited package", Assert.Single(reopened.Paragraphs).Text);
    }

    [Fact]
    public void FlatOpcPreservesInheritedMarkupCompatibilityPrefixAndRelativeMainPartTarget() {
        string xml = FlatOpc()
            .Replace("xmlPackage\">", "xmlPackage\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\">")
            .Replace("Target=\"word/document.xml\"", "Target=\"./word/document.xml\"")
            .Replace("2006/main\">", "2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"w14\">");
        using var source = new MemoryStream(Encoding.UTF8.GetBytes(xml));
        using WordDocument document = WordDocument.Load(source);
        using Stream part = document.OpenXmlDocument.MainDocumentPart.GetStream();
        XElement root = XElement.Load(part);
        XNamespace mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";
        Assert.Equal("w14", (string?)root.Attribute(mc + "Ignorable"));
        Assert.Equal("http://schemas.microsoft.com/office/word/2010/wordml", root.GetNamespaceOfPrefix("w14")?.NamespaceName);
        Assert.Equal("Producer package", Assert.Single(document.Paragraphs).Text);
    }

    [Fact]
    public void WordMl2003UsesSourceByteBudgetAndPreservesBreakClearing() {
        const string xml = """
            <w:wordDocument xmlns:w="http://schemas.microsoft.com/office/word/2003/wordml">
            <w:body><w:p><w:r><w:t>Before</w:t><w:br w:type="text-wrapping" w:clear="all"/><w:t>After</w:t></w:r></w:p></w:body>
            </w:wordDocument>
            """;
        byte[] bytes = Encoding.UTF8.GetBytes(xml);
        using var source = new MemoryStream(bytes);
        using WordDocument document = WordDocument.Load(source, new WordLoadOptions {
            FailOnLoss = true,
            MaxInputBytes = bytes.Length,
            PackageSecurity = new OfficePackageSecurityOptions { MaxPackageBytes = bytes.Length }
        });
        Break lineBreak = Assert.Single(document.OpenXmlDocument.MainDocumentPart.Document.Descendants<Break>());
        Assert.Equal(BreakTextRestartLocationValues.All, lineBreak.Clear!.Value);
    }

    [Fact]
    public void WordMl2003ProjectsBlocksAndReportsLossBeforeStrictConversion() {
        const string wordMl = """
            <?xml version="1.0"?>
            <?mso-application progid="Word.Document"?>
            <w:wordDocument xmlns:w="http://schemas.microsoft.com/office/word/2003/wordml"
              xmlns:wx="http://schemas.microsoft.com/office/word/2003/auxHint" xml:space="preserve">
              <w:body>
                <wx:sect>
                  <wx:sub-section>
                    <w:p><w:r><w:rPr><w:b/></w:rPr><w:t>  Report </w:t></w:r><w:hlink w:dest="https://example.invalid/"><w:r><w:t>visible link</w:t></w:r></w:hlink></w:p>
                    <w:tbl><w:tr><w:tc><w:p><w:r><w:t>Left</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Right</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
                  </wx:sub-section>
                  <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                </wx:sect>
              </w:body>
            </w:wordDocument>
            """;
        byte[] input = Encoding.UTF8.GetBytes(wordMl);
        using var source = new MemoryStream(input);
        using WordDocument document = WordDocument.Load(source, new WordLoadOptions(), out OfficeConversionReport report);
        Assert.Equal(WordFileFormat.WordMl2003, document.SourceFormat);
        Assert.Contains(report.Diagnostics, diagnostic =>
            diagnostic.LossKind == OfficeConversionLossKind.Omission && diagnostic.Location != null && diagnostic.Location.EndsWith("}hlink", StringComparison.Ordinal));

        using var output = new MemoryStream(document.ToBytes());
        using WordprocessingDocument saved = WordprocessingDocument.Open(output, false);
        Body body = saved.MainDocumentPart.Document.Body;
        Paragraph title = Assert.IsType<Paragraph>(body.FirstChild);
        Assert.Equal("  Report visible link", title.InnerText);
        Assert.NotNull(title.GetFirstChild<Run>().RunProperties.Bold);
        Assert.Equal(new[] { "Left", "Right" }, Assert.Single(body.Elements<Table>()).Descendants<TableCell>().Select(cell => cell.InnerText));
        Assert.Equal(12240U, body.GetFirstChild<SectionProperties>().GetFirstChild<PageSize>().Width.Value);

        using var strictSource = new MemoryStream(input);
        Assert.Throws<OfficeConversionException>(() => WordDocument.Load(strictSource,
            new WordLoadOptions { FailOnLoss = true, ReportUnsupportedContent = false }));
    }

    [Theory]
    [InlineData("duplicate")]
    [InlineData("spreadsheet")]
    [InlineData("depth")]
    public void InvalidFlatOpcFailsAtWordLoad(string boundary) {
        string xml = boundary switch {
            "duplicate" => FlatOpc().Replace("</pkg:package>", """
                  <pkg:part pkg:name="/WORD/document.xml" pkg:contentType="application/xml"><pkg:xmlData><other/></pkg:xmlData></pkg:part>
                </pkg:package>
                """),
            "spreadsheet" => FlatOpc().Replace(
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"),
            _ => FlatOpc().Replace("Producer package", string.Concat(Enumerable.Repeat("<nested>", 130)) +
                string.Concat(Enumerable.Repeat("</nested>", 130)))
        };
        using var source = new MemoryStream(Encoding.UTF8.GetBytes(xml));
        InvalidDataException error = Assert.Throws<InvalidDataException>(() => WordDocument.Load(source));
        Assert.Contains(boundary switch {
            "duplicate" => "duplicate",
            "spreadsheet" => "Word main document part",
            _ => "MaxDepth"
        }, error.Message);
    }

    private static string FlatOpc() => """
        <?xml version="1.0"?>
        <?mso-application progid="Word.Document"?>
        <pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
          <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
            <pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
            </Relationships></pkg:xmlData>
          </pkg:part>
          <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
            <pkg:xmlData><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Producer package</w:t></w:r></w:p></w:body></w:document></pkg:xmlData>
          </pkg:part>
        </pkg:package>
        """;
}
