using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Html;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests;

public partial class HtmlWordGapClosure {
    [Theory]
    [InlineData("""<html style="font-size:32px"><body><div style="margin-inline-start:1rem">Text</div></body></html>""")]
    [InlineData("""<style>html { font-size:32px; }</style><div style="margin-inline-start:1rem">Text</div>""")]
    public void HtmlToWord_RemLogicalSpacingUsesTheComputedRootFontSize(string html) {
        using WordDocument document = HtmlConversionDocument.Parse(html).ToWordDocument();
        WordParagraph paragraph = Assert.Single(
            document.Paragraphs,
            candidate => candidate.Text == "Text");

        Assert.Equal(480, paragraph.IndentationBefore);
    }

    [Fact]
    public void HtmlToWord_EmLogicalSpacingUsesTheComputedElementFontSize() {
        const string html = """
            <div style="font-size:20px;margin-inline-start:1em">Text</div>
            """;

        using WordDocument document = HtmlConversionDocument.Parse(html).ToWordDocument();
        WordParagraph paragraph = Assert.Single(
            document.Paragraphs,
            candidate => candidate.Text == "Text");

        Assert.Equal(300, paragraph.IndentationBefore);
    }

    [Fact]
    public void HtmlToWord_InheritedEmLogicalSpacingUsesTheParentsComputedFontSizeOnce() {
        const string html = """
            <style>.parent { font-size:2em; }</style>
            <div class="parent">
              <div style="margin-inline-start:1em">Text</div>
            </div>
            """;

        using WordDocument document = HtmlConversionDocument.Parse(html).ToWordDocument();
        WordParagraph paragraph = Assert.Single(
            document.Paragraphs,
            candidate => candidate.Text == "Text");

        Assert.Equal(480, paragraph.IndentationBefore);
    }

    [Fact]
    public void HtmlToWord_TopLevelInlineCodeStaysInTheBodyParagraph() {
        using WordDocument document = HtmlConversionDocument
            .Parse("""before<code>x</code>after""")
            .ToWordDocument();

        WordParagraph before = Assert.Single(document.Paragraphs, paragraph => paragraph.Text == "before");
        WordParagraph code = Assert.Single(document.Paragraphs, paragraph => paragraph.Text == "x");
        WordParagraph after = Assert.Single(document.Paragraphs, paragraph => paragraph.Text == "after");

        Assert.Same(before._paragraph, code._paragraph);
        Assert.Same(code._paragraph, after._paragraph);
    }

    [Fact]
    public void HtmlToWord_TableContainerEndSpacingConstrainsFullWidthTables() {
        const string html = """
            <div style="padding-inline-end:20px">
              <table style="width:100%"><tr><td>Value</td></tr></table>
            </div>
            """;

        using WordDocument document = HtmlConversionDocument.Parse(html).ToWordDocument();
        WordTable table = Assert.Single(document.Tables);
        WordSection section = Assert.Single(document.Sections);
        int pageWidth = (int)(section.PageSettings.Width ?? WordPageSizes.A4.WidthTwips);
        int contentWidth = pageWidth - (int)section.Margins.Left - (int)section.Margins.Right;

        Assert.Equal(WordTableWidthUnit.Dxa, table.WidthType);
        Assert.Equal(contentWidth - 300, table.Width);
    }

    [Fact]
    public void HtmlToWord_EmLogicalSpacingUsesTheFontShorthandSize() {
        const string html = """
            <div style="font:32px Arial;margin-inline-start:1em">Text</div>
            """;

        using WordDocument document = HtmlConversionDocument.Parse(html).ToWordDocument();
        WordParagraph paragraph = Assert.Single(
            document.Paragraphs,
            candidate => candidate.Text == "Text");

        Assert.Equal(480, paragraph.IndentationBefore);
    }

    [Theory]
    [InlineData("font-size:20px;font:32px Arial", 480)]
    [InlineData("font:32px Arial;font-size:20px", 300)]
    [InlineData("font-size:20px !important;font:32px Arial", 300)]
    public void HtmlToWord_FontSizeLonghandAndShorthandFollowTheCascade(
        string style,
        int expectedIndentation) {
        string html = $"""<div style="{style};margin-inline-start:1em">Text</div>""";

        using WordDocument document = HtmlConversionDocument.Parse(html).ToWordDocument();
        WordParagraph paragraph = Assert.Single(
            document.Paragraphs,
            candidate => candidate.Text == "Text");

        Assert.Equal(expectedIndentation, paragraph.IndentationBefore);
    }
}
