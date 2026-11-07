using System.Xml.Linq;
using OfficeIMO.Reader;
using OfficeIMO.Reader.IWork;
using OfficeIMO.Word;

namespace OfficeIMO.IWork.Tests;

public sealed class IWorkLegacyXmlTests {
    [Fact]
    public void NativePages09BodyOpensInReaderAndSurvivesAnEditableWordConversion() {
        string path = Path.Combine(AppContext.BaseDirectory, "Documents", "IWorkCorpus",
            "openpreserve", "lorem-ipsum.pages");
        using var input = File.OpenRead(path);
        using var source = IWorkSourceDocument.Load(input);
        Assert.Equal(IWorkDocumentKind.Pages, source.Kind);
        Assert.Empty(source.Records);
        Assert.Equal("Variatio Ipsius", source.ReadPages().Paragraphs.First(paragraph => paragraph.Length > 0));
        Assert.Throws<NotSupportedException>(() => source.BeginEdit());

        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddIWorkHandler().Build();
        OfficeDocumentReadResult read = reader.ReadDocument(input, "legacy.pages");
        Assert.Equal(ReaderInputKind.IWork, read.Kind);
        Assert.Contains("Lorem ipsum dolor sit amet", read.Markdown);
        Assert.True(input.CanRead);

        var conversion = new IWorkConversionOptions {
            Mode = IWorkConversionMode.EditableOnly,
            AllowPartialEditableReconstruction = true
        };
        using var result = source.ToWordResult(conversion);
        Assert.Contains(result.Report.Diagnostics, item => item.Code == "IWORK_LEGACY_PAGES_CONTENT_LIMITED"
            && item.LossKind == global::OfficeIMO.OfficeConversionLossKind.Omission);
        using var output = new MemoryStream();
        result.Value.Save(output);
        output.Position = 0;
        using WordDocument reopened = WordDocument.Load(output);
        Assert.Equal("Variatio Ipsius", reopened.Paragraphs.First(paragraph => paragraph.Text.Length > 0).Text);
        Assert.Contains(reopened.Paragraphs, paragraph => paragraph.Text.StartsWith("Lorem ipsum dolor sit amet",
            StringComparison.Ordinal));
    }

    [Fact]
    public void LegacyXmlIndexRejectsDtdBeforeProjectingBody() {
        using MemoryStream package = CreatePackage("""
            <!DOCTYPE sl:document [<!ENTITY unsafe "resolved">]>
            <sl:document xmlns:sl="http://developer.apple.com/namespaces/sl"
              xmlns:sf="http://developer.apple.com/namespaces/sf" sl:version="92008102400">
              <sf:text-storage><sf:text-body><sf:p>&unsafe;</sf:p></sf:text-body></sf:text-storage>
            </sl:document>
            """);
        var error = Assert.Throws<InvalidDataException>(() => IWorkSourceDocument.Load(package));
        Assert.IsType<System.Xml.XmlException>(error.InnerException);
        Assert.True(package.CanRead);
    }

    [Fact]
    public void LegacyXmlIndexRejectsExcessiveNestingDuringMaterialization() {
        string opening = string.Concat(Enumerable.Repeat("<sf:section>", 65));
        string closing = string.Concat(Enumerable.Repeat("</sf:section>", 65));
        using MemoryStream package = CreatePackage($"""
            <sl:document xmlns:sl="http://developer.apple.com/namespaces/sl"
              xmlns:sf="http://developer.apple.com/namespaces/sf" sl:version="92008102400">
              <sf:text-storage><sf:text-body>{opening}<sf:p>Text</sf:p>{closing}</sf:text-body></sf:text-storage>
            </sl:document>
            """);
        var error = Assert.Throws<InvalidDataException>(() => IWorkSourceDocument.Load(package));
        Assert.Contains("MaxDepth", error.Message);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LegacyXmlInlineControlsSurviveEditableConversion(bool keynote) {
        const string body = """
            <sf:text-storage><sf:text-body><sf:p>A<sf:tab/>B<sf:span><sf:lnbr/>C</sf:span><sf:crbr/>D<sf:intratopicbr/>E<sf:br/></sf:p></sf:text-body></sf:text-storage>
            """;
        string index = keynote ? $"""
            <key:presentation xmlns:key="http://developer.apple.com/namespaces/keynote2"
              xmlns:sf="http://developer.apple.com/namespaces/sf" xmlns:sfa="http://developer.apple.com/namespaces/sfa"
              key:version="92008102400">
              <key:slide-list><key:slide><key:title-placeholder sfa:ID="title"><key:text>{body}</key:text></key:title-placeholder>
              <key:page><sf:drawables><sf:title-placeholder-ref sfa:IDREF="title"/></sf:drawables></key:page>
              </key:slide></key:slide-list>
            </key:presentation>
            """ : $"""
            <sl:document xmlns:sl="http://developer.apple.com/namespaces/sl"
              xmlns:sf="http://developer.apple.com/namespaces/sf" sl:version="92008102400">{body}</sl:document>
            """;
        using MemoryStream package = CreatePackage(index, keynote ? "index.apxl" : "index.xml");
        using IWorkSourceDocument source = IWorkSourceDocument.Load(package);
        const string expected = "A\tB\nC\nD\nE";
        var options = new IWorkConversionOptions {
            Mode = IWorkConversionMode.EditableOnly,
            AllowPartialEditableReconstruction = true
        };
        using var output = new MemoryStream();
        if (keynote) {
            Assert.Equal(expected, Assert.Single(source.ReadKeynote().Slides).Title);
            using var result = source.ToPowerPointResult(options);
            result.Value.Save(output);
        } else {
            Assert.Equal(expected, Assert.Single(source.ReadPages().Paragraphs));
            using var result = source.ToWordResult(options);
            result.Value.Save(output);
        }
        output.Position = 0;
        using var archive = new ZipArchive(output, ZipArchiveMode.Read);
        using Stream payload = archive.GetEntry(keynote ? "ppt/slides/slide1.xml" : "word/document.xml")!.Open();
        XDocument document = XDocument.Load(payload, LoadOptions.PreserveWhitespace);
        XNamespace ns = keynote
            ? "http://schemas.openxmlformats.org/drawingml/2006/main"
            : "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        XElement paragraph = Assert.Single(document.Descendants(ns + "p"),
            item => item.Descendants(ns + "t").Any(text => text.Value == "A"));
        Assert.Equal(3, paragraph.Descendants(ns + "br").Count());
        if (keynote) {
            Assert.Equal("A\tBCDE", string.Concat(paragraph.Descendants(ns + "t").Select(text => text.Value)));
        } else {
            Assert.Single(paragraph.Descendants(ns + "tab"));
            Assert.Equal("ABCDE", string.Concat(paragraph.Descendants(ns + "t").Select(text => text.Value)));
        }
    }

    [Fact]
    public void LegacyKeynoteSlideNamesConsumeTheProjectionTextBudget() {
        using MemoryStream package = CreatePackage("""
            <key:presentation xmlns:key="http://developer.apple.com/namespaces/keynote2" key:version="92008102400">
              <key:slide-list><key:slide key:name="Native slide name"/></key:slide-list>
            </key:presentation>
            """, "index.apxl");
        using IWorkSourceDocument source = IWorkSourceDocument.Load(package,
            new IWorkLoadOptions { MaxProjectedTextCharacters = 8 });
        Assert.Throws<InvalidDataException>(() => source.ReadKeynote());
    }

    private static MemoryStream CreatePackage(string index, string indexName = "index.xml") {
        var package = new MemoryStream();
        using (var zip = new ZipArchive(package, ZipArchiveMode.Create, leaveOpen: true))
        using (var writer = new StreamWriter(zip.CreateEntry(indexName).Open()))
            writer.Write(index);
        package.Position = 0;
        return package;
    }
}
