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

        using WordDocument document = HtmlConversionDocument.Parse(html).ToWord();
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
            .ToWord();
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
        document.Sections[0].PageSettings.Width = global::OfficeIMO.OfficeLength.Points(600D);
        WordSection target = document.AddSection();
        target.PageSettings.Width = global::OfficeIMO.OfficeLength.Points(450D);
        target.Margins.Left = global::OfficeIMO.OfficeLength.Points(50D);
        target.Margins.Right = global::OfficeIMO.OfficeLength.Points(50D);

        document.AddHtml(HtmlConversionDocument.Parse("""
            <div style="padding-inline-end:20px">
              <table style="width:100%"><tr><td>Value</td></tr></table>
            </div>
            """));

        WordTable table = Assert.Single(document.Tables);
        Assert.Equal(WordTableWidthKind.Length, table.Width?.Kind);
        Assert.Equal(6700, table.Width?.ToNativeValue());
    }
}
