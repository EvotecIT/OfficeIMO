using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Html;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests;

public partial class HtmlWordGapClosure {
    [Fact]
    public void HtmlToWord_BlockBorderCurrentColorResolvesInheritedCssWideColor() {
        const string html = """
            <div style="color:red">
              <div style="color:inherit;border:1px solid">Text</div>
            </div>
            """;

        using WordDocument document = HtmlConversionDocument.Parse(html).ToWordDocument();
        WordParagraph paragraph = Assert.Single(
            document.Paragraphs,
            candidate => candidate.Text == "Text");

        Assert.Equal("FF0000", paragraph.Borders.LeftColorHex);
        Assert.Equal("FF0000", paragraph.Borders.RightColorHex);
    }

    [Fact]
    public void HtmlToWord_TransparentMarkClearsItsDefaultHighlightAndRoundTrips() {
        using WordDocument document = HtmlConversionDocument
            .Parse("""<mark style="background-color:transparent">Text</mark>""")
            .ToWordDocument();
        WordParagraph run = Assert.Single(
            document.Paragraphs,
            candidate => candidate.Text == "Text");

        Assert.Equal(WordHighlightColor.None, run.Highlight);
        Assert.Equal(string.Empty, run.RunShadingFillColorHex);
        Assert.Contains(
            """<mark style="background-color:transparent">Text</mark>""",
            document.ToHtml());
    }

    [Fact]
    public void AddHtmlToBody_TableContainerWidthUsesTheOwningSection() {
        using WordDocument document = WordDocument.Create();
        document.Sections[0].PageSettings.Width = 12000;
        WordSection target = document.AddSection();
        target.PageSettings.Width = 9000;
        target.Margins.Left = 1000;
        target.Margins.Right = 1000;

        document.AddHtmlToBody(HtmlConversionDocument.Parse("""
            <div style="padding-inline-end:20px">
              <table style="width:100%"><tr><td>Value</td></tr></table>
            </div>
            """));

        WordTable table = Assert.Single(document.Tables);
        Assert.Equal(WordTableWidthUnit.Dxa, table.WidthType);
        Assert.Equal(6700, table.Width);
    }
}
