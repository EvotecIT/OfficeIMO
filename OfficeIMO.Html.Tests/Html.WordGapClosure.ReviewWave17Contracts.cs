using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Html;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests;

public partial class HtmlWordGapClosure {
    [Fact]
    public void HtmlToWord_InlineCodeStaysInTheSurroundingContainerParagraph() {
        const string html = """<div>before<code>x</code>after</div>""";

        using WordDocument document = HtmlConversionDocument.Parse(html).ToWord();
        WordParagraph before = Assert.Single(document.Paragraphs, paragraph => paragraph.Text == "before");
        WordParagraph code = Assert.Single(document.Paragraphs, paragraph => paragraph.Text == "x");
        WordParagraph after = Assert.Single(document.Paragraphs, paragraph => paragraph.Text == "after");

        Assert.Same(before._paragraph, code._paragraph);
        Assert.Same(code._paragraph, after._paragraph);
    }

    [Fact]
    public void HtmlToWord_ParagraphBackgroundAlphaUsesTheEffectiveBackdrop() {
        const string html = """
            <p style="background-color:rgba(255,0,0,0)">Transparent</p>
            <div style="background-color:#0000ff">
              <p style="background-color:rgba(255,0,0,0.5)">Blended</p>
            </div>
            """;

        using WordDocument document = HtmlConversionDocument.Parse(html).ToWord();
        WordParagraph transparent = Assert.Single(document.Paragraphs, paragraph => paragraph.Text == "Transparent");
        WordParagraph blended = Assert.Single(document.Paragraphs, paragraph => paragraph.Text == "Blended");

        Assert.Null(transparent.ShadingFillColor);
        Assert.Equal(OfficeColor.Parse("800080"), blended.ShadingFillColor);
    }

    [Fact]
    public void WordRunShading_ExactFillResetsPatternAndForegroundState() {
        using WordDocument document = WordDocument.Create();
        WordParagraph run = document.AddParagraph("Pattern");
        run.BackgroundColor = OfficeColor.Parse("0000FF");
        Shading shading = run._runProperties!.Shading!;
        shading.Val = ShadingPatternValues.Percent50;
        shading.Color = "FF0000";
        shading.ThemeColor = ThemeColorValues.Accent1;
        shading.ThemeTint = "44";
        shading.ThemeShade = "77";

        run.BackgroundColor = OfficeColor.Parse("ABCDEF");

        Assert.Equal(OfficeColor.Parse("ABCDEF"), run.BackgroundColor);
        Assert.Equal(ShadingPatternValues.Clear, shading.Val?.Value);
        Assert.Null(shading.Color);
        Assert.Null(shading.ThemeColor);
        Assert.Null(shading.ThemeTint);
        Assert.Null(shading.ThemeShade);
    }

    [Fact]
    public void HtmlToWord_ContainerSpacingAppliesOncePerPhysicalParagraph() {
        const string html = """<div style="margin-left:10px"><p>a<strong>b</strong></p></div>""";

        using WordDocument document = HtmlConversionDocument.Parse(html).ToWord();
        WordParagraph firstRun = Assert.Single(document.Paragraphs, paragraph => paragraph.Text == "a");
        WordParagraph secondRun = Assert.Single(document.Paragraphs, paragraph => paragraph.Text == "b");

        Assert.Same(firstRun._paragraph, secondRun._paragraph);
        Assert.Equal(150, checked((int?)(firstRun.IndentationBefore?.ToTwips())));
    }

    [Fact]
    public void HtmlToWord_ContainerBorderUsesTheContainersCurrentColor() {
        const string html = """
            <div style="color:red;border:1px solid"><span style="color:blue">Text</span></div>
            """;

        using WordDocument document = HtmlConversionDocument.Parse(html).ToWord();
        WordParagraph paragraph = Assert.Single(document.Paragraphs, candidate => candidate.Text == "Text");

        Assert.Equal(OfficeColor.Parse("FF0000"), paragraph.Borders.LeftColor);
        Assert.Equal(OfficeColor.Parse("FF0000"), paragraph.Borders.RightColor);
        Assert.Equal(OfficeColor.Parse("FF0000"), paragraph.Borders.TopColor);
        Assert.Equal(OfficeColor.Parse("FF0000"), paragraph.Borders.BottomColor);
    }
}
