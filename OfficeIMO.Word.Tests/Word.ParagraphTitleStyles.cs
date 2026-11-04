using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(WordParagraphStyles.Title, "Title")]
    [InlineData(WordParagraphStyles.Subtitle, "Subtitle")]
    public void BuiltInTitleStylesCanBeSelectedAndReadBackThroughTheTypedApi(WordParagraphStyles style, string styleId) {
        using var document = WordDocument.Create();
        document.AddParagraph("Quarterly results").Style = style;
        Assert.Equal(styleId, WordParagraphStyle.GetStyleDefinition(style)?.StyleId);
        using var saved = WordDocument.Load(new System.IO.MemoryStream(document.ToBytes()));
        var paragraph = Assert.Single(saved.Paragraphs);
        Assert.Equal(style, paragraph.Style);
        Assert.Empty(saved.ValidateDocument());
    }
}
