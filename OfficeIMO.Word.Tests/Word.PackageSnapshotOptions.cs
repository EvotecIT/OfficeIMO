using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Word;
using Xunit;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void SharedPackageCloning_AssessmentEnforcesStrictLazyPartCharacterLimit() {
        using var input = new MemoryStream(CreateStrictSharedCloneInput(true, propertiesPadding: 4096));
        using WordDocument document = WordDocument.Load(input, new WordLoadOptions {
            AccessMode = DocumentAccessMode.ReadOnly,
            OpenSettings = new OfficeOpenXmlLoadSettings { MaxCharactersInPart = 1024 }
        });
        ExtendedFilePropertiesPart properties = document.OpenXmlDocument.ExtendedFilePropertiesPart!;
        Assert.False(properties.IsRootElementLoaded);

        Exception failure = Assert.ThrowsAny<Exception>(() => document.AssessLegacyDocWrite());

        Assert.Contains("MaxCharacters", failure.ToString());
        Assert.False(properties.IsRootElementLoaded);
        AssertSharedData(document.OpenXmlDocument);
    }

    [Fact]
    public void SharedPackageCloning_AssessmentEnforcesTransitionalLazyPartCharacterLimit() {
        using var input = new MemoryStream();
        byte[] bytes = CreateSharedCloneInput(true);
        input.Write(bytes, 0, bytes.Length);
        input.Position = 0;
        using (WordprocessingDocument package = WordprocessingDocument.Open(input, true)) {
            FootnotesPart footnotes = package.MainDocumentPart!.AddNewPart<FootnotesPart>();
            using var xml = new MemoryStream(Encoding.UTF8.GetBytes(
                "<w:footnotes xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                new string(' ', 4096) + "<w:footnote w:id=\"1\"><w:p><w:r><w:t>Note</w:t></w:r></w:p></w:footnote></w:footnotes>"));
            footnotes.FeedData(xml);
        }
        input.Position = 0;
        using WordDocument target = WordDocument.Load(input, new WordLoadOptions {
            AccessMode = DocumentAccessMode.ReadOnly,
            OpenSettings = new OfficeOpenXmlLoadSettings { MaxCharactersInPart = 1024 }
        });
        FootnotesPart lazy = target.OpenXmlDocument.MainDocumentPart!.FootnotesPart!;
        Assert.False(lazy.IsRootElementLoaded);
        Exception failure = Assert.ThrowsAny<Exception>(() => target.AssessLegacyDocWrite());

        Assert.Contains("MaxCharacters", failure.ToString());
        Assert.False(lazy.IsRootElementLoaded);
        AssertSharedData(target.OpenXmlDocument);
    }

    [Theory]
    [InlineData(FileFormatVersions.Office2007, "FALLBACK")]
    [InlineData(FileFormatVersions.Office2010, "CHOICE")]
    public void SharedPackageCloning_ProcessesLazyPartsForTheSourceCompatibilityVersion(
        FileFormatVersions version, string expected) {
        using var input = new MemoryStream();
        using (WordprocessingDocument package = WordprocessingDocument.Create(input, WordprocessingDocumentType.Document)) {
            MainDocumentPart main = package.AddMainDocumentPart();
            main.Document = new W.Document(new W.Body());
            HeaderPart header = main.AddNewPart<HeaderPart>("header");
            using var xml = new MemoryStream(Encoding.UTF8.GetBytes(
                "<w:hdr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" " +
                "xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" " +
                "xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\">" +
                "<mc:AlternateContent><mc:Choice Requires=\"w14\"><w:p><w:r><w:t>CHOICE</w:t></w:r></w:p></mc:Choice>" +
                "<mc:Fallback><w:p><w:r><w:t>FALLBACK</w:t></w:r></w:p></mc:Fallback></mc:AlternateContent></w:hdr>"));
            header.FeedData(xml);
        }
        input.Position = 0;
        using WordprocessingDocument source = WordprocessingDocument.Open(input, true, new OpenSettings {
            AutoSave = false,
            MarkupCompatibilityProcessSettings = new MarkupCompatibilityProcessSettings(
                MarkupCompatibilityProcessMode.ProcessAllParts, version)
        });
        HeaderPart lazy = Assert.Single(source.MainDocumentPart!.HeaderParts);
        Assert.False(lazy.IsRootElementLoaded);
        using var snapshot = new MemoryStream();

        using (WordPackageSnapshot.Clone(source, snapshot)) { }

        Assert.False(lazy.IsRootElementLoaded);
        snapshot.Position = 0;
        using WordprocessingDocument output = WordprocessingDocument.Open(snapshot, false);
        using Stream savedHeader = Assert.Single(output.MainDocumentPart!.HeaderParts).GetStream();
        XDocument headerXml = XDocument.Load(savedHeader);
        XNamespace word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        Assert.Equal(new[] { expected }, headerXml.Descendants(word + "t").Select(text => text.Value));
        Assert.Empty(headerXml.Descendants(XName.Get("AlternateContent", "http://schemas.openxmlformats.org/markup-compatibility/2006")));
    }
}
