using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Html;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeIMO.Tests;

public partial class HtmlWordGapClosure {
    [Fact]
    public void HtmlToWord_DirDoesNotCreateSelectorVisibleStyleAttributes() {
        const string html = """
            <style>[style] span { color:red; }</style>
            <div dir="rtl"><span>Text</span></div>
            """;

        using WordDocument document = HtmlConversionDocument.Parse(html).ToWord();
        WordParagraph paragraph = Assert.Single(
            document.Paragraphs,
            candidate => candidate.Text == "Text");
        WordParagraph run = Assert.Single(paragraph.GetRuns());

        Assert.True(paragraph.BiDi);
        Assert.NotEqual("FF0000", run.ColorHex);
    }

    [Fact]
    public void HtmlToWord_NestedContainerSpacingAccumulatesForFullWidthTables() {
        const string html = """
            <div style="padding-inline-end:20px">
              <div style="padding-inline-end:20px">
                <table style="width:100%"><tr><td>Value</td></tr></table>
              </div>
            </div>
            """;

        using WordDocument document = HtmlConversionDocument.Parse(html).ToWord();
        WordTable table = Assert.Single(document.Tables);
        WordSection section = Assert.Single(document.Sections);
        int pageWidth = (int)(checked((uint?)(section.PageSettings.Width?.ToTwips())) ?? checked((uint)(WordPageSizes.A4.Width.ToTwips())));
        int contentWidth = pageWidth - (int)checked((uint)(section.Margins.Left.GetValueOrDefault().ToTwips())) - (int)checked((uint)(section.Margins.Right.GetValueOrDefault().ToTwips()));

        Assert.Equal(WordTableWidthUnit.Dxa, table.WidthType);
        Assert.Equal(contentWidth - 600, table.Width);
    }

    [Fact]
    public void WordToHtml_ExactMarkedTextShadingIsAppliedToTheMarkElement() {
        using WordDocument document = HtmlConversionDocument
            .Parse("""<mark style="background-color:#ff0000">Text</mark>""")
            .ToWord();

        string html = document.ToHtml(new WordHtmlExportOptions { IncludeRunHighlightStyles = true });

        Assert.Contains(
            """<mark style="background-color:#ff0000">Text</mark>""",
            html,
            StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain(
            """background-color:#ff0000"><mark""",
            html,
            StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void WordToHtml_ExactMarkedTextShadingBesideAnEquationIsAppliedToTheMarkElement() {
        using WordDocument document = WordDocument.Create();
        WordParagraph paragraph = document.AddParagraph();
        paragraph._paragraph.Append(
            new Run(
                new RunProperties(
                    new RunStyle { Val = "HtmlMarkedText" },
                    new Shading { Val = ShadingPatternValues.Clear, Fill = "FF0000" }),
                new Text("Text")),
            new M.OfficeMath(new M.Run(new M.Text("x"))));

        string html = document.ToHtml(new WordHtmlExportOptions { IncludeRunHighlightStyles = true });

        Assert.Contains(
            """<mark style="background-color:#ff0000">Text</mark>""",
            html,
            StringComparison.OrdinalIgnoreCase);
        Assert.Contains("<math", html, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("margin-left:20px;margin-left:auto")]
    [InlineData("margin:20px;margin:auto")]
    [InlineData("margin-inline-start:20px;margin-inline-start:auto")]
    public void HtmlToWord_AutoMarginsClearEarlierNumericValues(string style) {
        string html = $"""<div style="{style}">Text</div>""";

        using WordDocument document = HtmlConversionDocument.Parse(html).ToWord();
        WordParagraph paragraph = Assert.Single(
            document.Paragraphs,
            candidate => candidate.Text == "Text");

        Assert.Null(checked((int?)(paragraph.IndentationBefore?.ToTwips())));
        Assert.Null(checked((int?)(paragraph.IndentationAfter?.ToTwips())));
        Assert.Null(checked((int?)(paragraph.LineSpacingBefore?.ToTwips())));
        Assert.Null(checked((int?)(paragraph.LineSpacingAfter?.ToTwips())));
    }
}
