using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Html;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests;

public partial class HtmlWordGapClosure {
    [Fact]
    public void HtmlToWord_TopLevelInlineImageStaysInTheSurroundingTextParagraph() {
        string html = $"""
            <body>before<img src="data:image/png;base64,{ValidPng}" alt="Inline">after</body>
        """;

        using WordDocument document = HtmlConversionDocument.Parse(html).ToWordDocument();
        WordParagraph before = Assert.Single(
            document.Paragraphs,
            candidate => candidate.Text == "before");
        WordParagraph after = Assert.Single(
            document.Paragraphs,
            candidate => candidate.Text == "after");

        Assert.Same(before._paragraph, after._paragraph);
        Assert.Single(before._paragraph.Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>());
        Assert.Single(document.Images);
    }

    [Fact]
    public void HtmlToWord_InlineBackgroundAlphaIsPreservedThroughTransparencyOrCompositing() {
        const string html = """
            <p><span style="background-color:rgba(255,0,0,0)">transparent</span><span style="background-color:rgba(255,0,0,0.5)">half</span></p>
            """;

        using WordDocument document = HtmlConversionDocument.Parse(html).ToWordDocument();
        WordParagraph transparent = Assert.Single(
            document.Paragraphs,
            run => run.Text == "transparent");
        WordParagraph half = Assert.Single(
            document.Paragraphs,
            run => run.Text == "half");

        Assert.Same(transparent._paragraph, half._paragraph);
        Assert.Equal(string.Empty, transparent.RunShadingFillColorHex);
        Assert.Equal("FF8080", half.RunShadingFillColorHex);
    }

    [Fact]
    public void WordToHtml_PercentageRunShadingUsesTheVisibleBlendedColor() {
        using WordDocument document = WordDocument.Create();
        WordParagraph run = document.AddParagraph("Pattern");
        run.RunShadingFillColorHex = "0000FF";
        run._runProperties!.Shading!.Val = ShadingPatternValues.Percent50;
        run._runProperties.Shading.Color = "FF0000";

        string html = document.ToHtml(new WordToHtmlOptions { IncludeRunHighlightStyles = true });
        WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();
        OfficeIMO.Drawing.OfficeRichTextRun renderedRun = Assert.Single(
            snapshot.Drawing.Elements
                .OfType<OfficeIMO.Drawing.OfficeDrawingRichText>()
                .SelectMany(richText => richText.Runs),
            candidate => candidate.Text == "Pattern");

        Assert.Contains("background-color:#800080", html, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(OfficeIMO.Drawing.OfficeColor.FromRgb(128, 0, 128), renderedRun.BackgroundColor);
    }

    [Fact]
    public void HtmlToWord_LogicalSpacingConvertsEveryAcceptedAbsoluteUnit() {
        const string html = """
            <p style="margin-inline-start:1rem">rem</p>
            <p style="padding-block-start:1in">in</p>
            <p style="margin-inline-start:1cm">cm</p>
            <p style="margin-inline-start:1mm">mm</p>
            <p style="margin-inline-start:1pc">pc</p>
            <p style="margin-inline-start:1q">q</p>
            """;
        var options = new HtmlToWordOptions {
            UnsupportedCssHandling = HtmlUnsupportedCssHandling.Error,
        };

        using WordDocument document = HtmlConversionDocument.Parse(html).ToWordDocument(options);

        Assert.Equal(240, FindParagraph(document, "rem").IndentationBefore);
        Assert.Equal(1440, FindParagraph(document, "in").LineSpacingBefore);
        Assert.Equal(567, FindParagraph(document, "cm").IndentationBefore);
        Assert.Equal(57, FindParagraph(document, "mm").IndentationBefore);
        Assert.Equal(240, FindParagraph(document, "pc").IndentationBefore);
        Assert.Equal(14, FindParagraph(document, "q").IndentationBefore);
    }

    private static WordParagraph FindParagraph(WordDocument document, string text) =>
        Assert.Single(document.Paragraphs, paragraph => paragraph.Text == text);
}
