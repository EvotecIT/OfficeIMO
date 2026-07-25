using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void RunShadingFillColor_DisabledPatternIgnoresStaleFill() {
        using WordDocument document = WordDocument.Create();
        WordParagraph paragraph = document.AddParagraph("Disabled");
        paragraph.RunShadingFillColorHex = "FF0000";
        paragraph._runProperties!.Shading!.Val = ShadingPatternValues.Nil;

        Assert.Equal(string.Empty, paragraph.RunShadingFillColorHex);
        Assert.Null(paragraph.RunShadingFillColor);

        using MemoryStream stream = document.ToStream();
        stream.Position = 0;
        using WordDocument reloaded = WordDocument.Load(stream);
        WordParagraph reloadedParagraph = Assert.Single(
            reloaded.Paragraphs,
            candidate => candidate.Text == "Disabled");

        Assert.Equal(string.Empty, reloadedParagraph.RunShadingFillColorHex);
        Assert.Null(reloadedParagraph.RunShadingFillColor);

        reloadedParagraph.RunShadingFillColorHex = "00FF00";

        Assert.Equal("00FF00", reloadedParagraph.RunShadingFillColorHex);
        Assert.Equal(ShadingPatternValues.Clear, reloadedParagraph._runProperties!.Shading!.Val?.Value);
    }

    [Fact]
    public void RunShadingFillColor_AutomaticFillReturnsNull() {
        using WordDocument document = WordDocument.Create();
        WordParagraph paragraph = document.AddParagraph("Automatic");

        paragraph.RunShadingFillColorHex = "auto";

        Assert.Equal("AUTO", paragraph.RunShadingFillColorHex);
        Assert.Null(paragraph.RunShadingFillColor);

        using MemoryStream stream = document.ToStream();
        stream.Position = 0;
        using WordDocument reloaded = WordDocument.Load(stream);
        WordParagraph reloadedParagraph = Assert.Single(
            reloaded.Paragraphs,
            candidate => candidate.Text == "Automatic");

        Assert.Equal("AUTO", reloadedParagraph.RunShadingFillColorHex);
        Assert.Null(reloadedParagraph.RunShadingFillColor);
    }
}
