using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Html;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests;

public partial class HtmlWordGapClosure {
    [Theory]
    [InlineData("break-before:page", true)]
    [InlineData("break-after:page", false)]
    public void HtmlToWord_TableOnlyContainerPreservesPageBreakBoundary(string style, bool breakBefore) {
        string html = $"""
            <div style="{style}"><table><tr><td>Cell</td></tr></table></div>
            """;

        using WordDocument document = HtmlConversionDocument.Parse(html).ToWordDocument();
        OpenXmlElement[] blocks = document.OpenXmlDocument.MainDocumentPart!.Document.Body!
            .ChildElements
            .Where(element => element is Paragraph or Table)
            .ToArray();

        Assert.Equal(2, blocks.Length);
        Assert.IsType<Paragraph>(blocks[breakBefore ? 0 : 1]);
        Assert.IsType<Table>(blocks[breakBefore ? 1 : 0]);
        Break pageBreak = Assert.Single(blocks.OfType<Paragraph>().Single().Descendants<Break>());
        Assert.Equal(BreakValues.Page, pageBreak.Type?.Value);
    }

    [Fact]
    public void HtmlToWord_StrictCssAcceptsTransparentInlineBackground() {
        const string html = """<p><span style="background-color:transparent">Clear</span></p>""";
        var options = new HtmlToWordOptions {
            UnsupportedCssHandling = HtmlUnsupportedCssHandling.Error,
        };

        using WordDocument document = HtmlConversionDocument.Parse(html).ToWordDocument(options);
        WordParagraph run = Assert.Single(document.Paragraphs, paragraph => paragraph.Text == "Clear");

        Assert.Equal(string.Empty, run.RunShadingFillColorHex);
    }

    [Fact]
    public void HtmlToWord_NestedContainerKeepsTheDescendantBorder() {
        const string html = """
            <div style="border:1px solid red">
              <div style="border:2px solid blue">Nested</div>
            </div>
            """;

        using WordDocument document = HtmlConversionDocument.Parse(html).ToWordDocument();
        WordParagraph paragraph = Assert.Single(document.Paragraphs, candidate => candidate.Text == "Nested");

        Assert.Equal(WordBorderStyle.Single, paragraph.Borders.LeftStyle);
        Assert.Equal("0000FF", paragraph.Borders.LeftColorHex);
        Assert.Equal(12U, paragraph.Borders.LeftSize?.Value);
    }

    [Fact]
    public void WordTableCell_AddHtml_NumberedHeadingKeepsInlinePlaceholderContent() {
        using WordDocument document = WordDocument.Create();
        WordTableCell cell = document.AddTable(1, 1).Rows[0].Cells[0];
        var options = new HtmlToWordOptions { SupportsHeadingNumbering = true };

        cell.AddHtml(HtmlConversionDocument.Parse("Intro<h1>Heading</h1>"), options);

        Assert.Equal(
            new[] { "Intro", "Heading" },
            cell._tableCell.Elements<Paragraph>().Select(paragraph => paragraph.InnerText).ToArray());
        WordParagraph heading = Assert.Single(cell.Paragraphs, paragraph => paragraph.Text == "Heading");
        Assert.Equal(WordListStyle.Headings111, heading.ListStyle);
    }
}
