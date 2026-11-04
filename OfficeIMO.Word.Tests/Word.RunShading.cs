using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void RunShadingFillColor_ClearsThemeFillAttributes() {
        using WordDocument document = WordDocument.Create();
        WordParagraph paragraph = document.AddParagraph("Themed");
        paragraph.BackgroundColor = OfficeColor.Parse("FFFFFF");
        paragraph._runProperties!.Shading!.ThemeFill = ThemeColorValues.Accent1;
        paragraph._runProperties.Shading.ThemeFillTint = "33";
        paragraph._runProperties.Shading.ThemeFillShade = "66";

        paragraph.BackgroundColor = OfficeColor.Parse("ABCDEF");

        Assert.Equal(OfficeColor.Parse("ABCDEF"), paragraph.BackgroundColor);
        Assert.Equal("ABCDEF", paragraph._runProperties.Shading.Fill?.Value);
        Assert.Null(paragraph._runProperties.Shading.ThemeFill);
        Assert.Null(paragraph._runProperties.Shading.ThemeFillTint);
        Assert.Null(paragraph._runProperties.Shading.ThemeFillShade);
    }

    [Fact]
    public void RunShadingFillColor_DisabledPatternIgnoresStaleFill() {
        using WordDocument document = WordDocument.Create();
        WordParagraph paragraph = document.AddParagraph("Disabled");
        paragraph.BackgroundColor = OfficeColor.Parse("FF0000");
        paragraph._runProperties!.Shading!.Val = ShadingPatternValues.Nil;

        Assert.Null(paragraph.BackgroundColor);
        Assert.Null(paragraph.BackgroundColor);

        using MemoryStream stream = new System.IO.MemoryStream(document.ToBytes());
        stream.Position = 0;
        using WordDocument reloaded = WordDocument.Load(stream);
        WordParagraph reloadedParagraph = Assert.Single(
            reloaded.Paragraphs,
            candidate => candidate.Text == "Disabled");

        Assert.Null(reloadedParagraph.BackgroundColor);
        Assert.Null(reloadedParagraph.BackgroundColor);

        reloadedParagraph.BackgroundColor = OfficeColor.Parse("00FF00");

        Assert.Equal(OfficeColor.Parse("00FF00"), reloadedParagraph.BackgroundColor);
        Assert.Equal(ShadingPatternValues.Clear, reloadedParagraph._runProperties!.Shading!.Val?.Value);
    }

    [Fact]
    public void RunShadingFillColor_AutomaticFillReturnsNull() {
        using WordDocument document = WordDocument.Create();
        WordParagraph paragraph = document.AddParagraph("Automatic");

        paragraph.RunShadingFillValue = "auto";

        Assert.Equal("AUTO", paragraph.RunShadingFillValue);
        Assert.Null(paragraph.BackgroundColor);

        using MemoryStream stream = new System.IO.MemoryStream(document.ToBytes());
        stream.Position = 0;
        using WordDocument reloaded = WordDocument.Load(stream);
        WordParagraph reloadedParagraph = Assert.Single(
            reloaded.Paragraphs,
            candidate => candidate.Text == "Automatic");

        Assert.Equal("AUTO", reloadedParagraph.RunShadingFillValue);
        Assert.Null(reloadedParagraph.BackgroundColor);
    }
}
