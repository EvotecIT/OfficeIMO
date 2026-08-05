using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Html;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests;

public partial class HtmlWordGapClosure {
    [Fact]
    public void HtmlToWord_StylesheetBlockImageStartsItsOwnParagraph() {
        string html = $$"""
            <style>.hero { display:block; }</style>
            <body>before<img class="hero" src="data:image/png;base64,{{ValidPng}}" alt="Hero">after</body>
            """;

        using WordDocument document = HtmlConversionDocument.Parse(html).ToWordDocument();
        WordParagraph before = Assert.Single(document.Paragraphs, paragraph => paragraph.Text == "before");
        WordParagraph image = Assert.Single(
            document.Paragraphs,
            paragraph => paragraph._paragraph.Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>().Any());
        WordParagraph after = Assert.Single(document.Paragraphs, paragraph => paragraph.Text == "after");

        Assert.NotSame(before._paragraph, image._paragraph);
        Assert.NotSame(image._paragraph, after._paragraph);
        Assert.NotSame(before._paragraph, after._paragraph);
    }

    [Fact]
    public void HtmlToWord_TableOnlyContainerFramesTheTableWithoutAnEmptyParagraph() {
        const string html = """
            <div style="background-color:#abcdef;border:2px solid #123456">
              <table><tr><td>First</td><td>Second</td></tr></table>
            </div>
            """;

        using WordDocument document = HtmlConversionDocument.Parse(html).ToWordDocument();
        WordTable table = Assert.Single(document.Sections[0].Tables);
        WordTableCell first = table.Rows[0].Cells[0];
        WordTableCell last = table.Rows[0].Cells[1];

        Assert.Empty(document.Sections[0].Paragraphs);
        Assert.Equal("ABCDEF", first.ShadingFillColorHex);
        Assert.Equal("ABCDEF", last.ShadingFillColorHex);
        Assert.Equal(WordBorderStyle.Single, first.Borders.TopStyle);
        Assert.Equal(WordBorderStyle.Single, first.Borders.LeftStyle);
        Assert.Equal(WordBorderStyle.Single, last.Borders.TopStyle);
        Assert.Equal(WordBorderStyle.Single, last.Borders.RightStyle);
        Assert.Equal("123456", first.Borders.LeftColorHex);
        Assert.Equal("123456", last.Borders.RightColorHex);
    }
}
