using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.LegacyDoc.Model;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(1)]
    [InlineData(50)]
    [InlineData(100)]
    [InlineData(200)]
    [InlineData(600)]
    public void NativeDocCharacterScalePreservesAuthoredWidthAndTracking(int percentage) {
        using WordDocument source = WordDocument.Create();
        WordParagraph paragraph = source.AddParagraph("Width and tracking");
        paragraph.CharacterScale = percentage;
        paragraph.Spacing = -10;
        using WordDocument loaded = WordDocument.Load(new MemoryStream(source.ToBytes(WordFileFormat.Doc)));
        Assert.Equal(percentage, loaded.Paragraphs[0].CharacterScale);
        Assert.Equal(-10, loaded.Paragraphs[0].Spacing);
        using WordDocument docx = WordDocument.Load(new MemoryStream(loaded.ToBytes()));
        Assert.Equal(percentage, docx.Paragraphs[0].CharacterScale);
        Assert.Empty(docx.ValidateDocument());
    }

    [Fact]
    public void NativeDocCharacterScalePreservesTablesLinksAndSeparateStories() {
        using WordDocument source = WordDocument.Create();
        source.AddParagraph("Body").CharacterScale = 75;
        source.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0].AddText("Cell").CharacterScale = 100;
        source.HeaderDefaultOrCreate.AddParagraph("Header").CharacterScale = 125;
        source.FooterDefaultOrCreate.AddParagraph("Footer").CharacterScale = 150;
        source.AddParagraph("Reference").AddFootNote("Footnote").FootNote!.Paragraphs!
            .Single(item => item.Text == "Footnote").CharacterScale = 50;
        source.AddParagraph("Reference").AddEndNote("Endnote").EndNote!.Paragraphs!
            .Single(item => item.Text == "Endnote").CharacterScale = 200;
        source.AddParagraph("Prefix ").AddHyperLink("Link", new Uri("https://example.test/width")).CharacterScale = 175;
        using WordDocument loaded = WordDocument.Load(new MemoryStream(source.ToBytes(WordFileFormat.Doc)));
        Assert.Equal(75, loaded.Paragraphs.Single(item => item.Text == "Body").CharacterScale);
        Assert.Equal(100, loaded.Tables[0].Rows[0].Cells[0].Paragraphs.Single(item => item.Text == "Cell").CharacterScale);
        Assert.Equal(125, loaded.Sections[0].Header.Default!.Paragraphs.Single(item => item.Text == "Header").CharacterScale);
        Assert.Equal(150, loaded.Sections[0].Footer.Default!.Paragraphs.Single(item => item.Text == "Footer").CharacterScale);
        Assert.Equal(50, Assert.Single(loaded.FootNotes).Paragraphs!.Single(item => item.Text == "Footnote").CharacterScale);
        Assert.Equal(200, Assert.Single(loaded.EndNotes).Paragraphs!.Single(item => item.Text == "Endnote").CharacterScale);
        WordParagraph link = Assert.Single(loaded.Paragraphs, item => item.IsHyperLink);
        Assert.Equal(175, link.CharacterScale);
        Assert.Equal(new Uri("https://example.test/width"), link.Hyperlink!.Uri);
        Assert.Empty(loaded.ValidateDocument());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NativeDocCharacterScaleRetainsInheritedWidthAndAnExplicitNormalReset(bool documentDefault) {
        using WordDocument source = WordDocument.Create();
        Styles styles = source._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        if (documentDefault) {
            styles.DocDefaults ??= new DocDefaults();
            styles.DocDefaults.RunPropertiesDefault ??= new RunPropertiesDefault(new RunPropertiesBaseStyle());
            styles.DocDefaults.RunPropertiesDefault.RunPropertiesBaseStyle!.AddChild(new CharacterScale { Val = 150L }, true);
        } else {
            var style = new Style { Type = StyleValues.Paragraph, StyleId = "ScaledParagraph", CustomStyle = true };
            style.Append(new StyleName { Val = "Scaled Paragraph" });
            style.StyleRunProperties = new StyleRunProperties(new CharacterScale { Val = 150L });
            styles.Append(style);
        }
        WordParagraph inherited = source.AddParagraph("Inherited");
        WordParagraph reset = source.AddParagraph("Reset");
        if (!documentDefault) { inherited.SetStyleId("ScaledParagraph"); reset.SetStyleId("ScaledParagraph"); }
        reset.CharacterScale = 100;
        using WordDocument loaded = WordDocument.Load(new MemoryStream(source.ToBytes(WordFileFormat.Doc)));
        WordParagraph loadedInherited = loaded.Paragraphs.Single(item => item.Text == "Inherited");
        Assert.Null(loadedInherited.CharacterScale);
        Style loadedStyle = loaded._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!
            .Elements<Style>().Single(item => item.StyleId == (loadedInherited.StyleId ?? "Normal"));
        Assert.Equal(150L, loadedStyle.StyleRunProperties!.GetFirstChild<CharacterScale>()!.Val!.Value);
        Assert.Equal(100, loaded.Paragraphs.Single(item => item.Text == "Reset").CharacterScale);
        Assert.Empty(loaded.ValidateDocument());
    }

    [Fact]
    public void NativeDocCharacterScaleRetainsParagraphMarkFormattingInSchemaOrder() {
        using WordDocument source = WordDocument.Create();
        WordParagraph paragraph = source.AddParagraph("Formatted paragraph mark");
        ParagraphProperties properties = paragraph._paragraphProperties ?? paragraph._paragraph.PrependChild(new ParagraphProperties());
        var mark = new ParagraphMarkRunProperties();
        mark.AddChild(new CharacterScale { Val = 200L }, true);
        mark.AddChild(new FontSize { Val = "32" }, true);
        mark.AddChild(new Spacing { Val = -10 }, true);
        properties.AddChild(mark, true);
        using WordDocument loaded = WordDocument.Load(new MemoryStream(source.ToBytes(WordFileFormat.Doc)));
        ParagraphMarkRunProperties loadedMark = loaded.Paragraphs[0]._paragraphProperties!.GetFirstChild<ParagraphMarkRunProperties>()!;
        Assert.Equal(200L, loadedMark.GetFirstChild<CharacterScale>()!.Val!.Value);
        Assert.Equal(-10, loadedMark.GetFirstChild<Spacing>()!.Val!.Value);
        Assert.Empty(loaded.ValidateDocument());
    }

    [Theory]
    [InlineData(0L)]
    [InlineData(601L)]
    [InlineData(long.MaxValue)]
    public void NativeDocCharacterScaleRejectsInvalidRawWidthInsteadOfTruncating(long percentage) {
        using WordDocument source = WordDocument.Create();
        WordParagraph paragraph = source.AddParagraph("Invalid authored width");
        paragraph._run!.RunProperties ??= new RunProperties();
        paragraph._run.RunProperties.AddChild(new CharacterScale { Val = percentage }, true);
        NotSupportedException error = Assert.Throws<NotSupportedException>(() => source.ToBytes(WordFileFormat.Doc));
        Assert.Contains("1 through 600", error.Message);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(601)]
    [InlineData(65535)]
    public void NativeDocCharacterScaleIgnoresInvalidOperandsWithoutLosingAdjacentFormatting(int invalid) {
        byte[] bytes = { 0x52, 0x48, 250, 0, 0x52, 0x48, (byte)invalid, (byte)(invalid >> 8), 0x35, 0x08, 1 };
        LegacyDocCharacterFormat format = LegacyDocCharacterFormattingReader.ReadGrpprl(bytes, 0, bytes.Length, Array.Empty<string>());
        Assert.Equal(250, format.CharacterScalePercentage);
        Assert.True(format.Bold);
    }

    [Fact]
    public void NativeDocCharacterScaleDoesNotReadAnOperandOutsideTheDeclaredRecord() {
        byte[] bytes = { 0x52, 0x48, 200, 0 };
        LegacyDocCharacterFormat format = LegacyDocCharacterFormattingReader.ReadGrpprl(bytes, 0, 3, Array.Empty<string>());
        Assert.Null(format.CharacterScalePercentage);
        Assert.False(format.IsSpecified(LegacyDocCharacterFormatProperties.CharacterScale));
    }
}
