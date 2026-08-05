using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Html;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests;

public partial class HtmlWordGapClosure {
    [Fact]
    public void WordTableCell_AddHtml_TableOnlyContainerUsesTheTableBoundaryForFramesAndBreaks() {
        using WordDocument document = WordDocument.Create();
        WordTableCell cell = document.AddTable(1, 1).Rows[0].Cells[0];

        cell.AddHtml(HtmlConversionDocument.Parse("""
            <div style="break-before:page;border:1px solid #123456;background-color:#abcdef">
              <table><tr><td>Value</td></tr></table>
            </div>
            """));

        var elements = cell._tableCell.ChildElements
            .Where(element => element is Table or Paragraph)
            .ToArray();
        Assert.Equal(3, elements.Length);
        Paragraph breakParagraph = Assert.IsType<Paragraph>(elements[0]);
        Break pageBreak = Assert.Single(breakParagraph.Descendants<Break>());
        Assert.Equal(BreakValues.Page, pageBreak.Type?.Value);
        Assert.IsType<Table>(elements[1]);
        Paragraph trailing = Assert.IsType<Paragraph>(elements[2]);
        Assert.Null(trailing.ParagraphProperties?.GetFirstChild<ParagraphBorders>());
        Assert.Null(trailing.ParagraphProperties?.GetFirstChild<Shading>());

        WordTableCell nestedCell = Assert.Single(Assert.Single(cell.DirectNestedTables).Rows).Cells[0];
        Assert.Equal("ABCDEF", nestedCell.ShadingFillColorHex);
        Assert.Equal("123456", nestedCell.Borders.LeftColorHex);
        Assert.Equal("123456", nestedCell.Borders.RightColorHex);
        Assert.Equal("123456", nestedCell.Borders.TopColorHex);
        Assert.Equal("123456", nestedCell.Borders.BottomColorHex);
    }

    [Fact]
    public void WordToHtml_HyperlinkRunBackgroundsUseTheHyperlinksRunProperties() {
        using WordDocument document = WordDocument.Create();
        WordParagraph highlightedParagraph = document.AddParagraph();
        highlightedParagraph.AddHyperLink("Highlighted", new Uri("https://example.test/highlight"));
        WordParagraph highlightedRun = Assert.Single(highlightedParagraph.GetRuns());
        highlightedRun.Highlight = WordHighlightColor.Yellow;

        WordParagraph shadedParagraph = document.AddParagraph();
        shadedParagraph.AddHyperLink("Shaded", new Uri("https://example.test/shading"));
        WordParagraph shadedRun = Assert.Single(shadedParagraph.GetRuns());
        shadedRun.RunShadingFillColorHex = "ABCDEF";

        string html = document.ToHtml(new WordToHtmlOptions { IncludeRunHighlightStyles = true });

        Assert.Contains("background-color:#ffff00", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("background-color:#abcdef", html, StringComparison.OrdinalIgnoreCase);
    }
}
