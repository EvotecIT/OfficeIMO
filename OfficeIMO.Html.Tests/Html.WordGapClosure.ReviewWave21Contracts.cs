using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Html;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests;

public partial class HtmlWordGapClosure {
    [Theory]
    [InlineData("background-color:rgba(255,0,0,0.5)", "", "")]
    [InlineData("", "background-color:rgba(255,0,0,0.5)", "")]
    [InlineData("", "", "background-color:rgba(255,0,0,0.5)")]
    public void HtmlToWord_TableAlphaUsesThePaintedAncestorBackdrop(
        string tableStyle,
        string rowStyle,
        string cellStyle) {
        string html = $"""
            <div style="background-color:#0000ff">
              <table style="{tableStyle}">
                <tr style="{rowStyle}">
                  <td style="{cellStyle}">Value</td>
                </tr>
              </table>
            </div>
            """;

        using WordDocument document = HtmlConversionDocument.Parse(html).ToWordDocument();
        WordTableCell cell = Assert.Single(Assert.Single(document.Tables).Rows).Cells[0];

        Assert.Equal("800080", cell.ShadingFillColorHex);
    }

    [Fact]
    public void HtmlToWord_ResourceClassificationDoesNotRerunSelectorsAgainstSyntheticStyles() {
        string html = $$"""
            <style>
              img { display:inline; }
              [style] { height:1px; }
            </style>
            <img width="100" src="data:image/png;base64,{{ValidPng}}" alt="Sized">
            """;

        using WordDocument document = HtmlConversionDocument.Parse(html).ToWordDocument();
        WordImage image = Assert.Single(document.Images);

        Assert.Equal(100D, Math.Round(image.Width!.Value));
        Assert.Equal(100D, Math.Round(image.Height!.Value));
    }

    [Fact]
    public void HtmlToWord_MixedContainerFrameUsesOnlyTheOuterBlockEdges() {
        const string html = """
            <div style="border:1px solid #123456">
              <p>Lead</p>
              <table><tr><td>Middle</td></tr></table>
              <table><tr><td>Last</td></tr></table>
            </div>
            """;

        using WordDocument document = HtmlConversionDocument.Parse(html).ToWordDocument();
        WordParagraph lead = Assert.Single(
            document.Paragraphs,
            paragraph => paragraph.Text == "Lead");
        WordTableCell middle = Assert.Single(document.Tables[0].Rows).Cells[0];
        WordTableCell last = Assert.Single(document.Tables[1].Rows).Cells[0];

        Assert.Equal(WordBorderStyle.Single, lead.Borders.TopStyle);
        Assert.Null(lead.Borders.BottomStyle);
        Assert.Null(middle.Borders.TopStyle);
        Assert.Null(middle.Borders.BottomStyle);
        Assert.Null(last.Borders.TopStyle);
        Assert.Equal(WordBorderStyle.Single, last.Borders.BottomStyle);
        Assert.Equal("123456", lead.Borders.LeftColorHex);
        Assert.Equal("123456", middle.Borders.LeftColorHex);
        Assert.Equal("123456", last.Borders.RightColorHex);
    }

    [Fact]
    public void HtmlToWord_BlockquoteVerticalSpacingAppliesOnlyAtTheOuterBoundaries() {
        const string html = """
            <blockquote style="margin-block:10px">
              <p>One</p>
              <p>Two</p>
            </blockquote>
            """;

        using WordDocument document = HtmlConversionDocument.Parse(html).ToWordDocument();
        WordParagraph one = Assert.Single(
            document.Paragraphs,
            paragraph => paragraph.Text == "One");
        WordParagraph two = Assert.Single(
            document.Paragraphs,
            paragraph => paragraph.Text == "Two");

        Assert.Equal(150, one.LineSpacingBefore);
        Assert.Null(one.LineSpacingAfter);
        Assert.Null(two.LineSpacingBefore);
        Assert.Equal(150, two.LineSpacingAfter);
    }
}
