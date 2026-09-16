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
    public void SettingsDependentSectionEditCreatesMissingSettingsPart() {
        using var stream = CreateMinimalDocument();
        using WordDocument document = WordDocument.Load(stream);
        MainDocumentPart mainPart = document._wordprocessingDocument.MainDocumentPart!;

        Assert.Null(mainPart.DocumentSettingsPart);

        document.Sections[0].DifferentOddAndEvenPages = true;

        Assert.NotNull(mainPart.DocumentSettingsPart?.Settings?.GetFirstChild<EvenAndOddHeaders>());
    }

    [Fact]
    public void DisablingDifferentOddAndEvenPagesDoesNotCreateMissingParts() {
        using var stream = CreateMinimalDocument();
        using WordDocument document = WordDocument.Load(stream);
        MainDocumentPart mainPart = document._wordprocessingDocument.MainDocumentPart!;

        Assert.Null(mainPart.DocumentSettingsPart);
        Assert.Empty(mainPart.HeaderParts);
        Assert.Empty(mainPart.FooterParts);

        document.Sections[0].DifferentOddAndEvenPages = false;

        Assert.Null(mainPart.DocumentSettingsPart);
        Assert.Empty(mainPart.HeaderParts);
        Assert.Empty(mainPart.FooterParts);
        Assert.False(document.Sections[0].DifferentOddAndEvenPages);
    }

    [Fact]
    public void DisablingDifferentOddAndEvenPagesRemovesExistingSetting() {
        using var stream = CreateMinimalDocument();
        using WordDocument document = WordDocument.Load(stream);
        MainDocumentPart mainPart = document._wordprocessingDocument.MainDocumentPart!;

        document.Sections[0].DifferentOddAndEvenPages = true;
        Assert.NotNull(mainPart.DocumentSettingsPart?.Settings?.GetFirstChild<EvenAndOddHeaders>());

        document.Sections[0].DifferentOddAndEvenPages = false;

        Assert.Null(mainPart.DocumentSettingsPart?.Settings?.GetFirstChild<EvenAndOddHeaders>());
        Assert.False(document.Sections[0].DifferentOddAndEvenPages);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void DifferentOddAndEvenPagesRecognizesOneDistinctEvenStory(bool removeHeaderReference) {
        using var stream = CreateMinimalDocument();
        using WordDocument document = WordDocument.Load(stream);
        WordSection section = document.Sections[0];
        section.DifferentOddAndEvenPages = true;

        if (removeHeaderReference) {
            Assert.Single(section._sectionProperties.Elements<HeaderReference>(), reference =>
                reference.Type?.Value == HeaderFooterValues.Even).Remove();
        } else {
            Assert.Single(section._sectionProperties.Elements<FooterReference>(), reference =>
                reference.Type?.Value == HeaderFooterValues.Even).Remove();
        }

        Assert.True(section.DifferentOddAndEvenPages);
        Assert.True(Assert.Single(document.CreateInspectionSnapshot().Sections).DifferentOddAndEvenPages);
    }

    [Fact]
    public void DifferentOddAndEvenPagesRecognizesInheritedEvenStories() {
        using var stream = CreateMinimalDocument();
        using WordDocument document = WordDocument.Load(stream);
        WordSection firstSection = document.Sections[0];
        firstSection.DifferentOddAndEvenPages = true;
        firstSection.Header.Even!.AddParagraph("Inherited even header");
        firstSection.Footer.Even!.AddParagraph("Inherited even footer");

        WordSection inheritedSection = document.AddSection();

        Assert.DoesNotContain(inheritedSection._sectionProperties.Elements<HeaderReference>(), reference =>
            reference.Type?.Value == HeaderFooterValues.Even);
        Assert.DoesNotContain(inheritedSection._sectionProperties.Elements<FooterReference>(), reference =>
            reference.Type?.Value == HeaderFooterValues.Even);
        Assert.True(inheritedSection.DifferentOddAndEvenPages);
        WordSectionSnapshot inheritedSnapshot = document.CreateInspectionSnapshot().Sections[1];
        Assert.True(inheritedSnapshot.DifferentOddAndEvenPages);
        Assert.Equal("Inherited even header", Assert.Single(inheritedSnapshot.EvenHeader!.Paragraphs).Text);
        Assert.Equal("Inherited even footer", Assert.Single(inheritedSnapshot.EvenFooter!.Paragraphs).Text);

        EvenAndOddHeaders setting = document._wordprocessingDocument.MainDocumentPart!
            .DocumentSettingsPart!
            .Settings!
            .GetFirstChild<EvenAndOddHeaders>()!;
        setting.Val = false;

        Assert.False(firstSection.DifferentOddAndEvenPages);
        Assert.False(inheritedSection.DifferentOddAndEvenPages);
        WordSectionSnapshot disabledSnapshot = document.CreateInspectionSnapshot().Sections[1];
        Assert.False(disabledSnapshot.DifferentOddAndEvenPages);
        Assert.Equal("Inherited even header", Assert.Single(disabledSnapshot.EvenHeader!.Paragraphs).Text);
        Assert.Equal("Inherited even footer", Assert.Single(disabledSnapshot.EvenFooter!.Paragraphs).Text);

        firstSection.DifferentOddAndEvenPages = true;
        Assert.True(setting.Val!.Value);
        Assert.True(inheritedSection.DifferentOddAndEvenPages);

        using var reloadedStream = new MemoryStream(document.ToBytes(), writable: false);
        using WordDocument reloaded = WordDocument.Load(reloadedStream);
        Assert.True(reloaded.Sections[1].DifferentOddAndEvenPages);
        Assert.Equal("Inherited even header", Assert.Single(
            reloaded.CreateInspectionSnapshot().Sections[1].EvenHeader!.Paragraphs).Text);
    }

    [Fact]
    public void ParagraphOnOffPropertiesHonorExplicitFalseValues() {
        using var stream = CreateMinimalDocument();
        using WordDocument document = WordDocument.Load(stream);
        WordParagraph paragraph = Assert.Single(document.Paragraphs);
        ParagraphProperties properties = paragraph._paragraph.ParagraphProperties ??= new ParagraphProperties();
        properties.PageBreakBefore = new PageBreakBefore { Val = false };
        properties.KeepNext = new KeepNext { Val = false };
        properties.KeepLines = new KeepLines { Val = false };
        properties.WidowControl = new WidowControl { Val = false };
        properties.BiDi = new BiDi { Val = false };

        Assert.False(paragraph.PageBreakBefore);
        Assert.False(paragraph.KeepWithNext);
        Assert.False(paragraph.KeepLinesTogether);
        Assert.False(paragraph.AvoidWidowAndOrphan);
        Assert.False(paragraph.BiDi);

        WordParagraphSnapshot snapshot = Assert.IsType<WordParagraphSnapshot>(
            Assert.Single(Assert.Single(document.CreateInspectionSnapshot().Sections).Elements));
        Assert.False(snapshot.PageBreakBefore);
        Assert.False(snapshot.KeepWithNext);
        Assert.False(snapshot.KeepLinesTogether);
        Assert.False(snapshot.AvoidWidowAndOrphan);
        Assert.False(snapshot.IsRightToLeft);
    }

    [Fact]
    public void InspectionSnapshotPreservesTextHiddenByRunFormatting() {
        using var stream = CreateMinimalDocument();
        using WordDocument document = WordDocument.Load(stream);
        WordParagraph paragraph = Assert.Single(document.Paragraphs);
        Run run = Assert.Single(paragraph._paragraph.Elements<Run>());
        run.RunProperties = new RunProperties(new Vanish());

        WordParagraphSnapshot snapshot = Assert.IsType<WordParagraphSnapshot>(
            Assert.Single(Assert.Single(document.CreateInspectionSnapshot().Sections).Elements));

        Assert.Equal("Minimal document", snapshot.Text);
        Assert.Equal("Minimal document", Assert.Single(snapshot.Runs).Text);
    }

    [Fact]
    public void BackgroundImageEditCreatesMissingSettingsPartBeforeEditing() {
        using var stream = CreateMinimalDocument();
        using WordDocument document = WordDocument.Load(stream);
        MainDocumentPart mainPart = document._wordprocessingDocument.MainDocumentPart!;
        byte[] imageBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP4/w8AAv8B/h10yjMAAAAASUVORK5CYII=");
        using var imageStream = new MemoryStream(imageBytes, writable: false);

        Assert.Null(mainPart.DocumentSettingsPart);

        document.Background.SetImage(imageStream, "background.png", 1, 1);

        Assert.NotNull(mainPart.DocumentSettingsPart?.Settings?.DisplayBackgroundShape);
        Assert.NotNull(mainPart.Document?.DocumentBackground);
    }

    [Fact]
    public void StyleDependentSettingsEditCreatesMissingStylesPart() {
        using var stream = CreateMinimalDocument(includeStyles: false);
        using WordDocument document = WordDocument.Load(stream);
        MainDocumentPart mainPart = document._wordprocessingDocument.MainDocumentPart!;

        Assert.Null(mainPart.StyleDefinitionsPart);

        document.Settings.FontSize = 13;
        document.Settings.FontFamily = "Aptos";

        Assert.NotNull(mainPart.StyleDefinitionsPart?.Styles);
        Assert.Equal(13, document.Settings.FontSize);
        Assert.Equal("Aptos", document.Settings.FontFamily);
        Assert.Empty(new OpenXmlValidator().Validate(document._wordprocessingDocument));
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
    public void CompleteCatalogCacheDoesNotMaskASeparateIncompleteCatalog() {
        byte[] complete;
        using (WordDocument created = WordDocument.Create()) {
            created.AddParagraph("Complete catalog");
            complete = created.ToBytes();
        }

        for (int iteration = 0; iteration < 2; iteration++) {
            using var completeStream = new MemoryStream(complete, writable: false);
            using WordDocument completeDocument = WordDocument.Load(completeStream);
            complete = completeDocument.ToBytes();
        }

        using var minimalStream = CreateMinimalDocument();
        using WordDocument minimalDocument = WordDocument.Load(minimalStream);
        byte[] savedMinimal = minimalDocument.ToBytes();
        using var savedStream = new MemoryStream(savedMinimal, writable: false);
        using WordprocessingDocument savedPackage = WordprocessingDocument.Open(savedStream, false);

        Assert.Contains("TableGrid", GetStyleIds(savedPackage.MainDocumentPart!));
        Assert.Contains("Header", GetStyleIds(savedPackage.MainDocumentPart!));
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
