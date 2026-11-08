using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Drawing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordXmlInterchangeTests {
    [Theory]
    [InlineData(WordFileFormat.FlatOpc, false)]
    [InlineData(WordFileFormat.FlatOpc, true)]
    [InlineData(WordFileFormat.WordMl2003, false)]
    [InlineData(WordFileFormat.WordMl2003, true)]
    public async Task EncryptedXmlCopyRejectsXmlDestinationsWithoutWritingAndAcceptsPackageDestinations(WordFileFormat format, bool useAsync) {
        string xml = format == WordFileFormat.FlatOpc ? FlatOpc() : """
            <w:wordDocument xmlns:w="http://schemas.microsoft.com/office/word/2003/wordml">
              <w:body><w:p><w:r><w:t>Producer package</w:t></w:r></w:p></w:body>
            </w:wordDocument>
            """;
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-word-xml-copy-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            using var input = new MemoryStream(Encoding.UTF8.GetBytes(xml));
            using WordDocument document = WordDocument.Load(input);
            Assert.Equal(format, document.SourceFormat);
            Assert.Single(document.Paragraphs).SetText("Edited encrypted copy");
            const string password = "xml-copy-software-fixture";
            var options = new WordSaveOptions { Password = password };
            Task Copy(string destination) {
                if (useAsync) return document.SaveCopyAsync(destination, options);
                document.SaveCopy(destination, options);
                return Task.CompletedTask;
            }

            string existing = Path.Combine(directory, "separate.xml");
            byte[] sentinel = Encoding.UTF8.GetBytes("<sentinel>existing destination</sentinel>");
            File.WriteAllBytes(existing, sentinel);
            NotSupportedException existingError = await Assert.ThrowsAsync<NotSupportedException>(() => Copy(existing));
            Assert.Contains("non-XML destination", existingError.Message);
            Assert.Equal(sentinel, File.ReadAllBytes(existing));

            string missing = Path.Combine(directory, "new.xml");
            NotSupportedException missingError = await Assert.ThrowsAsync<NotSupportedException>(() => Copy(missing));
            Assert.Contains("non-XML destination", missingError.Message);
            Assert.False(File.Exists(missing));

            string package = Path.Combine(directory, "copy.docx");
            await Copy(package);
            using WordDocument reopened = WordDocument.Load(package, new WordLoadOptions { Password = password });
            Assert.Equal("Edited encrypted copy", Assert.Single(reopened.Paragraphs).Text);
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void WordMlSourceFormatDescribesMacroCapabilityWhileReportingDroppedCarrier() {
        const string xml = """
            <w:wordDocument xmlns:w="http://schemas.microsoft.com/office/word/2003/wordml" w:macrosPresent="on">
              <w:docSuppData><w:binData w:name="editdata.mso">AA==</w:binData></w:docSuppData>
              <w:body><w:p><w:r><w:t>Inert migration source</w:t></w:r></w:p></w:body>
            </w:wordDocument>
            """;
        using var input = new MemoryStream(Encoding.UTF8.GetBytes(xml));
        using WordDocument document = WordDocument.Load(input, null, out OfficeConversionReport report);
        Assert.True(document.SourceFormatDescriptor.IsMacroEnabled);
        Assert.False(document.HasMacros);
        Assert.Contains(report.Diagnostics, finding => finding.LossKind == OfficeConversionLossKind.Omission &&
            finding.Location?.EndsWith("}docSuppData", StringComparison.Ordinal) == true);
        Assert.Equal("Inert migration source", document.Paragraphs.Single().Text);
        input.Position = 0;
        Assert.Throws<OfficeConversionException>(() => WordDocument.Load(input, new WordLoadOptions { FailOnLoss = true }));
    }

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
              xmlns:wx="http://schemas.microsoft.com/office/word/2003/auxHint"
              xmlns:first="urn:word-source:first" xmlns:second="urn:word-source:second" xml:space="preserve">
              <w:body>
                <first:extra/><second:extra/><first:extra/>
                <wx:sect>
                  <wx:sub-section>
                    <w:p><w:pPr><w:framePr w:w="2000"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>  Report </w:t></w:r><w:hlink w:dest="https://example.invalid/"><w:r><w:t>visible link</w:t></w:r></w:hlink></w:p>
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
            diagnostic.LossKind == OfficeConversionLossKind.Omission && diagnostic.Location != null && diagnostic.Location.EndsWith("}framePr", StringComparison.Ordinal));

        using var output = new MemoryStream(document.ToBytes());
        using WordprocessingDocument saved = WordprocessingDocument.Open(output, false);
        Body body = saved.MainDocumentPart.Document.Body;
        Paragraph title = Assert.IsType<Paragraph>(body.FirstChild);
        Assert.Equal("  Report visible link", title.InnerText);
        Assert.NotNull(title.GetFirstChild<Run>().RunProperties.Bold);
        Assert.Equal("https://example.invalid/", saved.MainDocumentPart.HyperlinkRelationships.Single().Uri.AbsoluteUri);
        Assert.Equal(saved.MainDocumentPart.HyperlinkRelationships.Single().Id, title.GetFirstChild<Hyperlink>().Id.Value);
        Assert.Equal(new[] { "Left", "Right" }, Assert.Single(body.Elements<Table>()).Descendants<TableCell>().Select(cell => cell.InnerText));
        Assert.Equal(12240U, body.GetFirstChild<SectionProperties>().GetFirstChild<PageSize>().Width.Value);

        using var strictSource = new MemoryStream(input);
        Assert.Throws<OfficeConversionException>(() => WordDocument.Load(strictSource,
            new WordLoadOptions { FailOnLoss = true, ReportUnsupportedContent = false }));

        string inputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".docx");
        string outputPath = Path.ChangeExtension(inputPath, ".doc");
        try {
            File.WriteAllBytes(inputPath, input);
            var expectedLocations = new[] { "{urn:word-source:first}extra", "{urn:word-source:second}extra" };
            const string omissionMessage = "WordML 2003 block content was omitted (extra).";
            WordDocumentConversionReport analysis = WordDocument.AnalyzeConversion(inputPath, outputPath);
            OfficeConversionDiagnostic[] omissions = analysis.NativeDiagnostics
                .Where(finding => finding.Code == "WORDML2003_OMITTED" && finding.Message == omissionMessage).ToArray();
            Assert.Equal(expectedLocations, omissions.Select(finding => finding.SourceLocation));
            Assert.All(omissions, finding => {
                Assert.True(finding.RepresentsDataLoss);
                Assert.Equal(OfficeCompatibilityState.Blocked, finding.CompatibilityState);
            });
            Assert.Equal(expectedLocations, analysis.Diagnostics.Where(finding => finding.Message == omissionMessage)
                .Select(finding => finding.Location));
            Assert.False(File.Exists(outputPath));

            WordDocumentConversionException blocked = Assert.Throws<WordDocumentConversionException>(() =>
                WordDocument.Convert(inputPath, outputPath));
            Assert.Equal(OfficeConversionFailureReason.DataLossBlocked, blocked.Reason);
            Assert.Equal(expectedLocations, blocked.Result.Report.NativeDiagnostics
                .Where(finding => finding.Message == omissionMessage).Select(finding => finding.SourceLocation));
            Assert.False(File.Exists(outputPath));

            WordDocumentConversionReport allowed = WordDocument.AnalyzeConversion(inputPath, outputPath,
                new WordDocumentConversionOptions { LossPolicy = OfficeConversionLossPolicy.Allow });
            Assert.Equal(expectedLocations, allowed.NativeDiagnostics
                .Where(finding => finding.Message == omissionMessage).Select(finding => finding.SourceLocation));
            Assert.All(allowed.NativeDiagnostics.Where(finding => finding.Message == omissionMessage),
                finding => Assert.Equal(OfficeCompatibilityState.Dropped, finding.CompatibilityState));
            Assert.False(File.Exists(outputPath));
        } finally {
            File.Delete(inputPath);
            File.Delete(outputPath);
        }
    }

    [Theory]
    [InlineData("duplicate")]
    [InlineData("spreadsheet")]
    [InlineData("depth")]
    [InlineData("part-uri")]
    [InlineData("relationship-uri")]
    public void InvalidFlatOpcFailsAtWordLoad(string boundary) {
        string xml = boundary switch {
            "duplicate" => FlatOpc().Replace("</pkg:package>", """
                  <pkg:part pkg:name="/WORD/document.xml" pkg:contentType="application/xml"><pkg:xmlData><other/></pkg:xmlData></pkg:part>
                </pkg:package>
                """),
            "spreadsheet" => FlatOpc().Replace(
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"),
            "part-uri" => FlatOpc().Replace("pkg:name=\"/word/document.xml\"", "pkg:name=\"/word/\""),
            "relationship-uri" => FlatOpc().Replace("Target=\"word/document.xml\"", "Target=\"word/\""),
            _ => FlatOpc().Replace("Producer package", string.Concat(Enumerable.Repeat("<nested>", 130)) +
                string.Concat(Enumerable.Repeat("</nested>", 130)))
        };
        using var source = new MemoryStream(Encoding.UTF8.GetBytes(xml));
        InvalidDataException error = Assert.Throws<InvalidDataException>(() => WordDocument.Load(source));
        Assert.Contains(boundary switch {
            "duplicate" => "duplicate",
            "spreadsheet" => "Word main document part",
            "part-uri" => "part name",
            "relationship-uri" => "Word main",
            _ => "MaxDepth"
        }, error.Message);
    }

    [Theory]
    [InlineData(WordFileFormat.WordMl2003, "truncated")]
    [InlineData(WordFileFormat.WordMl2003, "second-root")]
    [InlineData(WordFileFormat.FlatOpc, "truncated")]
    [InlineData(WordFileFormat.FlatOpc, "second-root")]
    public async Task RecognizedWordXmlSyntaxErrorsUseConsistentAdmissionErrors(WordFileFormat format, string syntax) {
        string xml = format == WordFileFormat.FlatOpc ? FlatOpc() : """
            <w:wordDocument xmlns:w="http://schemas.microsoft.com/office/word/2003/wordml">
              <w:body><w:p><w:r><w:t>Producer package</w:t></w:r></w:p></w:body>
            </w:wordDocument>
            """;
        string closingTag = format == WordFileFormat.FlatOpc ? "</pkg:package>" : "</w:wordDocument>";
        xml = syntax == "truncated" ? xml.Replace(closingTag, "") : xml + "<unexpected/>";
        byte[] bytes = Encoding.UTF8.GetBytes(xml);

        using var source = new MemoryStream(bytes);
        InvalidDataException streamError = Assert.Throws<InvalidDataException>(() => WordDocument.Load(source));
        Assert.IsType<XmlException>(streamError.InnerException);
        Assert.True(source.CanRead);

        using var asyncSource = new MemoryStream(bytes);
        InvalidDataException asyncStreamError = await Assert.ThrowsAsync<InvalidDataException>(() => WordDocument.LoadAsync(asyncSource));
        Assert.IsType<XmlException>(asyncStreamError.InnerException);
        Assert.True(asyncSource.CanRead);

        string path = Path.Combine(Path.GetTempPath(), "officeimo-invalid-word-xml-" + Guid.NewGuid().ToString("N") + ".xml");
        File.WriteAllBytes(path, bytes);
        try {
            InvalidDataException fileError = Assert.Throws<InvalidDataException>(() => WordDocument.Load(path));
            Assert.IsType<XmlException>(fileError.InnerException);
            InvalidDataException asyncFileError = await Assert.ThrowsAsync<InvalidDataException>(() => WordDocument.LoadAsync(path));
            Assert.IsType<XmlException>(asyncFileError.InnerException);
        } finally {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("unknown-root")]
    [InlineData("unknown-truncated")]
    [InlineData("before-root")]
    public async Task UnrecognizedXmlKeepsNormalPackageAdmissionErrors(string boundary) {
        string xml = boundary switch {
            "unknown-root" => "<unknown/>",
            "unknown-truncated" => "<unknown>",
            _ => "<?xml version=\"1.0\"?><"
        };
        byte[] bytes = Encoding.UTF8.GetBytes(xml);
        using var source = new MemoryStream(bytes);
        FileFormatException error = Assert.Throws<FileFormatException>(() => WordDocument.Load(source));
        Assert.Null(error.InnerException);
        using var asyncSource = new MemoryStream(bytes);
        FileFormatException asyncError = await Assert.ThrowsAsync<FileFormatException>(() => WordDocument.LoadAsync(asyncSource));
        Assert.Null(asyncError.InnerException);
    }

    [Theory]
    [InlineData("AB!D", null)]
    [InlineData("A=BC", null)]
    [InlineData("AA=A", null)]
    [InlineData("QUJD", "ABC")]
    [InlineData("QUI=", "AB")]
    [InlineData(" Q Q = = \r\n\t", "A")]
    public void FlatOpcBinaryPayloadsUseConsistentAdmissionErrorsAndPreserveValidBytes(string payload, string? expected) {
        string xml = FlatOpc().Replace("</pkg:package>", """
          <pkg:part pkg:name="/word/_rels/document.xml.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
            <pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rPayload" Type="urn:officeimo:test:payload" Target="custom.bin"/>
            </Relationships></pkg:xmlData>
          </pkg:part>
          <pkg:part pkg:name="/word/custom.bin" pkg:contentType="application/octet-stream"><pkg:binaryData>
          """ + payload + "</pkg:binaryData></pkg:part></pkg:package>");
        using var source = new MemoryStream(Encoding.UTF8.GetBytes(xml));
        if (expected == null) {
            InvalidDataException error = Assert.Throws<InvalidDataException>(() => WordDocument.Load(source));
            Assert.Contains("base64", error.Message);
            Assert.IsType<FormatException>(error.InnerException);
            Assert.True(source.CanRead);
            return;
        }

        using WordDocument document = WordDocument.Load(source, new WordLoadOptions { FailOnLoss = true }, out OfficeConversionReport report);
        Assert.False(report.HasLoss);
        using var bytes = new MemoryStream(document.ToBytes());
        using WordprocessingDocument saved = WordprocessingDocument.Open(bytes, false);
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Microsoft365).Validate(saved));
        Assert.Equal("Producer package", saved.MainDocumentPart!.Document.InnerText);
        ExtendedPart part = Assert.Single(saved.MainDocumentPart.GetPartsOfType<ExtendedPart>());
        using Stream content = part.GetStream(FileMode.Open, FileAccess.Read);
        using var actual = new MemoryStream();
        content.CopyTo(actual);
        Assert.Equal(Encoding.ASCII.GetBytes(expected), actual.ToArray());
        Assert.True(source.CanRead);
    }

    [Theory]
    [InlineData("<w:sectPr/>")]
    [InlineData("<w:p><w:pPr><w:sectPr/></w:pPr></w:p>")]
    public void WordMlReportsSectionPropertiesOmittedInsideTableCells(string block) {
        string xml = "<w:wordDocument xmlns:w=\"http://schemas.microsoft.com/office/word/2003/wordml\">" +
            "<w:body><w:tbl><w:tr><w:tc>" + block + "</w:tc></w:tr></w:tbl></w:body></w:wordDocument>";
        byte[] bytes = Encoding.UTF8.GetBytes(xml);
        using var source = new MemoryStream(bytes);
        using WordDocument document = WordDocument.Load(source, new WordLoadOptions(), out OfficeConversionReport report);
        Assert.Contains(report.Diagnostics, diagnostic => diagnostic.LossKind == OfficeConversionLossKind.Omission &&
            diagnostic.Location?.EndsWith("}sectPr", StringComparison.Ordinal) == true);
        Assert.Throws<OfficeConversionException>(() => WordDocument.Load(source, new WordLoadOptions { FailOnLoss = true }));
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
