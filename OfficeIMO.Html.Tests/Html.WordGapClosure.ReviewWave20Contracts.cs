using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Html;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests;

public partial class HtmlWordGapClosure {
    [Fact]
    public void HtmlToWord_TableOnlyContainerAppliesItsBoxSpacingToTheTableBoundary() {
        const string html = """
            <div style="margin-top:10px;padding-left:20px">
              <table><tr><td>Value</td></tr></table>
            </div>
            """;

        using WordDocument document = HtmlConversionDocument.Parse(html).ToWordDocument();
        WordTable table = Assert.Single(document.Tables);
        Paragraph spacing = Assert.IsType<Paragraph>(
            table._table.PreviousSibling());

        Assert.Equal((short)300, table.StyleDetails?.TableIndentationWidth);
        Assert.Equal("150", spacing.ParagraphProperties?
            .GetFirstChild<SpacingBetweenLines>()?.After?.Value);
    }

    [Fact]
    public void HtmlToWord_InlineAlphaCompositesAgainstTheParagraphBackground() {
        const string html = """
            <p style="background-color:#0000ff">
              plain <span style="background-color:rgba(255,0,0,0.5)">blended</span>
            </p>
            """;

        using WordDocument document = HtmlConversionDocument.Parse(html).ToWordDocument();
        WordParagraph paragraph = Assert.Single(
            document.Paragraphs,
            candidate => candidate.Text == "blended");

        Assert.Equal("0000FF", paragraph.ShadingFillColorHex);
        Assert.Equal("800080", paragraph.RunShadingFillColorHex);
    }

    [Fact]
    public void HtmlToWord_BodyLogicalSpacingIsDiagnosedAsUnsupported() {
        HtmlUnsupportedCssException exception = Assert.Throws<HtmlUnsupportedCssException>(() =>
            HtmlConversionDocument
                .Parse("""<body style="margin:0;margin-inline-start:10px"><p>Text</p></body>""")
                .ToWordDocument(new HtmlToWordOptions {
                    UnsupportedCssHandling = HtmlUnsupportedCssHandling.Error
                }));

        Assert.Equal("UnsupportedCssDeclaration", exception.Code);
        Assert.Equal("body:margin-inline-start", exception.CssSource);
    }

    [Fact]
    public void HtmlToWord_TransparentAndAlphaTableBackgroundsUseTheExistingBackdrop() {
        const string html = """
            <table style="background-color:#0000ff">
              <tr>
                <td style="background-color:transparent">transparent</td>
                <td style="background-color:rgba(255,0,0,0.5)">blended</td>
              </tr>
            </table>
            """;

        using WordDocument document = HtmlConversionDocument.Parse(html).ToWordDocument();
        List<WordTableCell> cells = Assert.Single(document.Tables).Rows[0].Cells;

        Assert.Equal("0000FF", cells[0].ShadingFillColorHex);
        Assert.Equal("800080", cells[1].ShadingFillColorHex);
    }

    [Fact]
    public void WordTableCell_AddHtml_BelowCaptionClearsSyntheticDirectSpacing() {
        using WordDocument document = WordDocument.Create();
        WordTableCell cell = document.AddTable(1, 1).Rows[0].Cells[0];

        cell.AddHtml(
            HtmlConversionDocument.Parse(
                """<table><caption>Nested caption</caption><tr><td>Value</td></tr></table>"""),
            new HtmlToWordOptions { TableCaptionPosition = TableCaptionPosition.Below });

        Paragraph caption = Assert.Single(
            cell._tableCell.Elements<Paragraph>(),
            paragraph => paragraph.InnerText == "Nested caption");
        Assert.Equal("Caption", caption.ParagraphProperties?.ParagraphStyleId?.Val?.Value);
        Assert.Null(caption.ParagraphProperties?.GetFirstChild<SpacingBetweenLines>());
    }

    [Fact]
    public void HtmlToWord_MixedContainerBreaksUseTheTrueContentBoundaries() {
        using WordDocument tableFirst = HtmlConversionDocument
            .Parse("""
                <div style="break-before:page">
                  <table><tr><td>Value</td></tr></table>
                  <p>Tail</p>
                </div>
                """)
            .ToWordDocument();
        OpenXmlElement[] tableFirstBlocks = GetBodyBlocks(tableFirst);
        int tableIndex = Array.FindIndex(tableFirstBlocks, element => element is Table);
        int tailIndex = Array.FindIndex(
            tableFirstBlocks,
            element => element is Paragraph paragraph && paragraph.InnerText == "Tail");
        int breakBeforeIndex = Array.FindIndex(
            tableFirstBlocks,
            element => element.Descendants<Break>().Any());

        Assert.True(breakBeforeIndex < tableIndex);
        Assert.True(tableIndex < tailIndex);

        using WordDocument tableLast = HtmlConversionDocument
            .Parse("""
                <div style="break-after:page">
                  <p>Lead</p>
                  <table><tr><td>Value</td></tr></table>
                </div>
                """)
            .ToWordDocument();
        OpenXmlElement[] tableLastBlocks = GetBodyBlocks(tableLast);
        int leadIndex = Array.FindIndex(
            tableLastBlocks,
            element => element is Paragraph paragraph && paragraph.InnerText == "Lead");
        int lastTableIndex = Array.FindIndex(tableLastBlocks, element => element is Table);
        int breakAfterIndex = Array.FindLastIndex(
            tableLastBlocks,
            element => element.Descendants<Break>().Any());

        Assert.True(leadIndex < lastTableIndex);
        Assert.True(lastTableIndex < breakAfterIndex);
    }

    private static OpenXmlElement[] GetBodyBlocks(WordDocument document) =>
        document.OpenXmlDocument.MainDocumentPart!.Document.Body!.ChildElements
            .Where(element => element is Paragraph or Table)
            .ToArray();
}
