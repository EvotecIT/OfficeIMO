using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(WordFileFormat.Doc)]
    [InlineData(WordFileFormat.Docx)]
    public void LegacyDoc_ExplicitAutomaticRunColorOverridesInheritedRgb(WordFileFormat format) {
        using WordDocument document = WordDocument.Create();
        Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        Style normal = styles.Elements<Style>().Single(style => style.StyleId?.Value == "Normal");
        normal.StyleRunProperties ??= new StyleRunProperties();
        normal.StyleRunProperties.Color = new Color { Val = "FF0000" };
        document.AddParagraph("Automatic").RunColorValue = "auto";

        using WordDocument loaded = WordDocument.Load(new MemoryStream(document.ToBytes(new WordSaveOptions { Format = format })));
        WordParagraph actual = Assert.Single(loaded.Paragraphs, paragraph => paragraph.Text == "Automatic");
        Assert.Equal("auto", actual.RunColorValue);
        Assert.Null(actual.Color);
        Assert.Empty(loaded.DocumentValidationErrors);
    }

    [Fact]
    public void LegacyDoc_ParagraphStyleAutomaticRunColorSurvivesNativeSave() {
        using WordDocument document = WordDocument.Create();
        Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        styles.Append(new Style(new StyleName { Val = "Automatic color style" },
            new BasedOn { Val = "Normal" }, new StyleRunProperties(new Color { Val = "auto" })) {
            StyleId = "AutomaticColorStyle", Type = StyleValues.Paragraph, CustomStyle = true
        });
        document.AddParagraph("Styled").SetStyleId("AutomaticColorStyle");

        using WordDocument loaded = WordDocument.Load(new MemoryStream(document.ToBytes(new WordSaveOptions { Format = WordFileFormat.Doc })));
        Style actual = loaded._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!
            .Elements<Style>().Single(style => style.StyleName?.Val?.Value == "Automatic color style");
        Assert.Equal("auto", actual.StyleRunProperties?.Color?.Val?.Value);
        Assert.Empty(loaded.DocumentValidationErrors);
    }

    [Fact]
    public void LegacyDoc_ImportsAutomaticRunColorFromMicrosoftWordAndResavesIt() {
        using WordDocument source = WordDocument.Load(GetFixtureDoc(Path.Combine("RunColors", "word-automatic-font-colors.doc")));
        WordParagraph automatic = Assert.Single(source.Paragraphs, paragraph => paragraph.Text == "Auto000000");
        // Word omits a direct declaration when the automatic document default applies.
        Assert.Null(automatic.Color);
        automatic.RunColorValue = "auto";
        Assert.Equal(OfficeColor.Parse("000000"), Assert.Single(source.Paragraphs, paragraph => paragraph.Text == "ExplicitBlack010203").Color);

        using WordDocument loaded = WordDocument.Load(new MemoryStream(source.ToBytes(new WordSaveOptions { Format = WordFileFormat.Doc })));
        Assert.Equal("auto", Assert.Single(loaded.Paragraphs, paragraph => paragraph.Text == "Auto000000").RunColorValue);
        Assert.Equal(OfficeColor.Parse("000000"), Assert.Single(loaded.Paragraphs, paragraph => paragraph.Text == "ExplicitBlack010203").Color);
        Assert.Empty(loaded.DocumentValidationErrors);
    }
}
