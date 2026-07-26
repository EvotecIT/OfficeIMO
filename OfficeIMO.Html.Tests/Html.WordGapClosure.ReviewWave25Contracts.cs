using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Html;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeIMO.Tests;

public partial class HtmlWordGapClosure {
    [Fact]
    public void HtmlToWord_TransparentContainersPreserveTheEffectiveBackdrop() {
        const string html = """
            <div style="background-color:#0000ff">
              <div style="background-color:transparent">
                <div style="background-color:rgba(255,0,0,0.5)">Text</div>
              </div>
            </div>
            """;

        using WordDocument document = HtmlConversionDocument.Parse(html).ToWordDocument();
        WordParagraph paragraph = Assert.Single(
            document.Paragraphs,
            candidate => candidate.Text == "Text");

        Assert.Equal("800080", paragraph.ShadingFillColorHex);
    }

    [Fact]
    public void WordToHtml_NonDefaultMarkedTextHighlightIsAppliedToTheMarkElement() {
        using WordDocument document = WordDocument.Create();
        document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Append(
            new Style(
                new StyleName { Val = "Marked text" },
                new StyleRunProperties(new Highlight { Val = HighlightColorValues.Green })) {
                Type = StyleValues.Character,
                StyleId = "HtmlMarkedText"
            });
        WordParagraph regular = document.AddParagraph("Regular");
        regular.SetCharacterStyleId("HtmlMarkedText");
        WordParagraph equationAdjacent = document.AddParagraph();
        equationAdjacent._paragraph.Append(
            new Run(
                new RunProperties(new RunStyle { Val = "HtmlMarkedText" }),
                new Text("Adjacent")),
            new M.OfficeMath(new M.Run(new M.Text("x"))));

        string html = document.ToHtml(new WordToHtmlOptions { IncludeRunHighlightStyles = true });

        Assert.Contains(
            """<mark style="background-color:#00ff00">Regular</mark>""",
            html,
            StringComparison.OrdinalIgnoreCase);
        Assert.Contains(
            """<mark style="background-color:#00ff00">Adjacent</mark>""",
            html,
            StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void WordTableCell_AddHtml_PercentageImageUsesTheCellContentWidth() {
        using WordDocument document = WordDocument.Create();
        WordTableCell cell = document.AddTable(1, 2).Rows[0].Cells[0];
        int expectedContentWidth = Assert.IsType<int>(
            WordTable.EstimateCellContentWidthInDxa(document, cell._tableCell));
        WordSection laterSection = document.AddSection();
        laterSection.PageSettings.Width = 20000;
        laterSection.Margins.Left = 500;
        laterSection.Margins.Right = 500;

        cell.AddHtml(HtmlConversionDocument.Parse(
            $"""<img style="width:100%" src="data:image/png;base64,{ValidPng}" alt="Cell image">"""));

        WordImage image = Assert.Single(document.Images);
        Assert.Equal(expectedContentWidth / 15D, image.Width!.Value, precision: 3);
    }

    [Fact]
    public void HtmlToWord_ModernRgbBackgroundSyntaxPreservesColorAndAlpha() {
        using WordDocument document = HtmlConversionDocument
            .Parse("""<div style="background-color:rgb(100% 0% 50% / 0.5)">Text</div>""")
            .ToWordDocument();
        WordParagraph paragraph = Assert.Single(
            document.Paragraphs,
            candidate => candidate.Text == "Text");

        Assert.Equal("FF80C0", paragraph.ShadingFillColorHex);
    }

    [Theory]
    [InlineData("border:solid red")]
    [InlineData("border-left-style:solid;border-left-color:red")]
    public void HtmlToWord_OmittedBlockBorderWidthUsesCssMedium(string style) {
        string html = $"""<div style="{style}">Text</div>""";

        using WordDocument document = HtmlConversionDocument.Parse(html).ToWordDocument();
        WordParagraph paragraph = Assert.Single(
            document.Paragraphs,
            candidate => candidate.Text == "Text");

        Assert.Equal(BorderValues.Single, paragraph.Borders.LeftStyle);
        Assert.Equal(18U, paragraph.Borders.LeftSize?.Value);
    }
}
