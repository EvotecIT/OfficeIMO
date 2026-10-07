using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void NativeDocCharacterScalePreservesStandardHeaderFooterStyles(bool footer, bool directReset) {
        using WordDocument source = WordDocument.Create();
        source.AddParagraph("Body");
        WordParagraph paragraph = footer ? source.FooterDefaultOrCreate.AddParagraph("Story width")
            : source.HeaderDefaultOrCreate.AddParagraph("Story width");
        string styleId = footer ? "Footer" : "Header";
        paragraph.SetStyleId(styleId);
        Styles styles = source._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        Style style = styles.Elements<Style>().Single(item => item.StyleId == styleId);
        style.StyleRunProperties ??= new StyleRunProperties();
        style.StyleRunProperties.AddChild(new CharacterScale { Val = 200L }, true);
        style.StyleRunProperties.AddChild(new Spacing { Val = 20 }, true);
        if (directReset) { paragraph.CharacterScale = 100; paragraph.Spacing = 0; }
        string before = style.OuterXml;
        Assert.Empty(source.ValidateDocument());

        using WordDocument loaded = WordDocument.Load(new MemoryStream(source.ToBytes(WordFileFormat.Doc)));
        WordParagraph actual = (footer ? loaded.Sections[0].Footer.Default!.Paragraphs : loaded.Sections[0].Header.Default!.Paragraphs)
            .Single(item => item.Text == "Story width");
        Style imported = loaded._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!
            .Elements<Style>().Single(item => item.StyleId == actual.StyleId);
        Assert.Equal(200L, imported.StyleRunProperties?.CharacterScale?.Val?.Value);
        Assert.Equal(20, imported.StyleRunProperties?.Spacing?.Val?.Value);
        Assert.Equal(directReset ? 100 : (int?)null, actual.CharacterScale);
        Assert.Equal(directReset ? 0 : (int?)null, actual.Spacing);
        Assert.Equal(before, style.OuterXml);
        Assert.Empty(loaded.ValidateDocument());
    }
}
