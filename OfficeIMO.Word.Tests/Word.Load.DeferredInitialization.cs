using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordLoadDeferredInitializationTests {
    [Fact]
    public void ReadingDocumentDoesNotAddUnrelatedPackageParts() {
        using var stream = CreateMinimalDocument();
        using WordDocument document = WordDocument.Load(stream);

        WordprocessingDocument package = document._wordprocessingDocument;
        MainDocumentPart mainPart = package.MainDocumentPart!;
        Assert.Null(mainPart.DocumentSettingsPart);
        Assert.Equal(new[] { "Normal" }, GetStyleIds(mainPart));

        WordParagraph paragraph = Assert.Single(document.Paragraphs);
        Assert.Equal("Minimal document", paragraph.Text);

        Assert.Null(mainPart.DocumentSettingsPart);
        Assert.Equal(new[] { "Normal" }, GetStyleIds(mainPart));
    }

    [Fact]
    public void AccessingSettingsCreatesMissingSettingsPart() {
        using var stream = CreateMinimalDocument();
        using WordDocument document = WordDocument.Load(stream);
        MainDocumentPart mainPart = document._wordprocessingDocument.MainDocumentPart!;

        Assert.Null(mainPart.DocumentSettingsPart);

        Assert.NotNull(document.Settings);
        Assert.NotNull(mainPart.DocumentSettingsPart?.Settings);
    }

    [Fact]
    public void SavingLoadedDocumentAddsRequiredStyleCatalog() {
        using var stream = CreateMinimalDocument();
        using WordDocument document = WordDocument.Load(stream);

        byte[] saved = document.ToBytes();

        using var savedStream = new MemoryStream(saved, writable: false);
        using WordprocessingDocument package = WordprocessingDocument.Open(savedStream, false);
        MainDocumentPart mainPart = package.MainDocumentPart!;
        string[] styleIds = GetStyleIds(mainPart);
        Assert.Contains("Normal", styleIds);
        Assert.Contains("TableGrid", styleIds);
        Assert.Contains("Header", styleIds);
        Assert.Null(mainPart.DocumentSettingsPart);
        Assert.Equal("Minimal document", mainPart.Document.Body!.InnerText);
        Assert.Empty(new OpenXmlValidator().Validate(package));
    }

    [Fact]
    public void SavingStylelessDocumentUsesAvailableRelationshipId() {
        using var stream = CreateMinimalDocument(includeStyles: false, settingsUsesFirstRelationshipId: true);
        using WordDocument document = WordDocument.Load(stream);

        byte[] saved = document.ToBytes();

        using var savedStream = new MemoryStream(saved, writable: false);
        using WordprocessingDocument package = WordprocessingDocument.Open(savedStream, false);
        MainDocumentPart mainPart = package.MainDocumentPart!;
        Assert.NotNull(mainPart.DocumentSettingsPart);
        Assert.Contains("TableGrid", GetStyleIds(mainPart));
        Assert.Empty(new OpenXmlValidator().Validate(package));
    }

    [Fact]
    public void PlainParagraphFastPathRemainsEditableAndRoundTrips() {
        using WordDocument document = WordDocument.Create();
        WordParagraph paragraph = document.AddParagraph("  plain text  ").SetBold();

        Assert.Equal("  plain text  ", paragraph.Text);
        Assert.True(paragraph.Bold);

        using var savedStream = new MemoryStream(document.ToBytes(), writable: false);
        using WordDocument reloaded = WordDocument.Load(savedStream);
        WordParagraph reloadedParagraph = Assert.Single(reloaded.Paragraphs);
        Assert.Equal("  plain text  ", reloadedParagraph.Text);
        Assert.True(reloadedParagraph.Bold);
    }

    [Fact]
    public void NullParagraphTextRetainsEmptyParagraphCompatibility() {
        using WordDocument document = WordDocument.Create();

        WordParagraph paragraph = document.AddParagraph((string)null!);

        Assert.Equal(string.Empty, paragraph.Text);
        Assert.Single(document.Paragraphs);
    }

    private static MemoryStream CreateMinimalDocument(
        bool includeStyles = true,
        bool settingsUsesFirstRelationshipId = false) {
        var stream = new MemoryStream();
        using (WordprocessingDocument package = WordprocessingDocument.Create(
                   stream,
                   WordprocessingDocumentType.Document,
                   autoSave: true)) {
            MainDocumentPart mainPart = package.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                new Paragraph(new Run(new Text("Minimal document")))));

            if (settingsUsesFirstRelationshipId) {
                DocumentSettingsPart settingsPart = mainPart.AddNewPart<DocumentSettingsPart>("rId1");
                settingsPart.Settings = new Settings();
            }

            if (includeStyles) {
                StyleDefinitionsPart stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                stylesPart.Styles = new Styles(
                    new Style(
                        new StyleName { Val = "Normal" }) {
                        Type = StyleValues.Paragraph,
                        StyleId = "Normal",
                        Default = true
                    });
            }
        }

        stream.Position = 0;
        return stream;
    }

    private static string[] GetStyleIds(MainDocumentPart mainPart) {
        return mainPart.StyleDefinitionsPart!.Styles!
            .Elements<Style>()
            .Select(style => style.StyleId?.Value)
            .Where(styleId => styleId != null)
            .Cast<string>()
            .ToArray();
    }
}
