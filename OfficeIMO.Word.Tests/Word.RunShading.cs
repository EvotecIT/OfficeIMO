using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
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
