using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Html;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests;

public partial class HtmlWordGapClosure {
    [Fact]
    public void HtmlToWord_ContainerFramePreservesExistingTableCellBorders() {
        const string html = """
            <div style="border:2px solid red">
              <table><tr><td style="border:1px solid blue">Cell</td></tr></table>
            </div>
            """;

        using WordDocument document = HtmlConversionDocument.Parse(html).ToWordDocument();
        WordTableCell cell = Assert.Single(Assert.Single(document.Tables).Rows).Cells[0];

        Assert.Equal("0000FF", cell.Borders.LeftColorHex);
        Assert.Equal("0000FF", cell.Borders.RightColorHex);
        Assert.Equal("0000FF", cell.Borders.TopColorHex);
        Assert.Equal("0000FF", cell.Borders.BottomColorHex);
    }

    [Fact]
    public void WordToHtml_CharacterStyleHighlightTakesPrecedenceOverRunShading() {
        using WordDocument document = WordDocument.Create();
        document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Append(
            new Style(
                new StyleName { Val = "Styled highlight" },
                new StyleRunProperties(new Highlight { Val = HighlightColorValues.Yellow })) {
                Type = StyleValues.Character,
                StyleId = "StyledHighlight"
            });
        WordParagraph run = document.AddParagraph("Layered");
        run.SetCharacterStyleId("StyledHighlight");
        run.RunShadingFillColorHex = "FF0000";

        string html = document.ToHtml(new WordToHtmlOptions { IncludeRunHighlightStyles = true });

        Assert.Contains("background-color:#ffff00", html, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("background-color:#ff0000", html, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void HtmlToWord_BlockBorderAlphaIsPreservedVisually() {
        const string html = """
            <div style="background-color:#0000ff;border-left-style:solid;border-left-width:1px;border-left-color:rgba(255,0,0,0)">Transparent</div>
            <div style="background-color:#0000ff;border-left-style:solid;border-left-width:1px;border-left-color:rgba(255,0,0,0.5)">Blended</div>
            """;

        using WordDocument document = HtmlConversionDocument.Parse(html).ToWordDocument();
        WordParagraph transparent = Assert.Single(document.Paragraphs, paragraph => paragraph.Text == "Transparent");
        WordParagraph blended = Assert.Single(document.Paragraphs, paragraph => paragraph.Text == "Blended");

        Assert.Null(transparent.Borders.LeftStyle);
        Assert.Equal(BorderValues.Single, blended.Borders.LeftStyle);
        Assert.Equal("800080", blended.Borders.LeftColorHex);
    }

    [Fact]
    public void HtmlToWord_MarginOnlyEmptyBlockMaterializesSpacing() {
        const string html = """
            <p>Before</p>
            <div style="margin-block:20px"></div>
            <p>After</p>
            """;

        using WordDocument document = HtmlConversionDocument.Parse(html).ToWordDocument();
        WordParagraph spacing = Assert.Single(
            document.Paragraphs,
            paragraph => string.IsNullOrEmpty(paragraph.Text));

        Assert.Equal(300, spacing.LineSpacingBefore);
        Assert.Equal(300, spacing.LineSpacingAfter);
    }
}
