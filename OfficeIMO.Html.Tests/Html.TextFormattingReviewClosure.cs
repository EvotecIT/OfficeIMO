using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.OneNote;
using OfficeIMO.OneNote.Html;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlTextFormattingReviewClosureTests {
    [Fact]
    public void PowerPointSemanticHtmlRoundTripPreservesParagraphBreaksAndRunFormatting() {
        using PowerPointPresentation source = PowerPointPresentation.Create();
        PowerPointTextBox textBox = source.AddSlide().AddTextBox("First");
        textBox.Paragraphs[0].Runs[0].Bold = true;
        PowerPointParagraph second = textBox.AddParagraph("Second");
        second.Runs[0].Italic = true;

        string html = source.ToHtml(PowerPointHtmlSaveOptions.CreateSemanticSlidesProfile());
        using PowerPointPresentation imported = HtmlConversionDocument
            .Parse(html, HtmlConversionDocumentOptions.CreateTrustedProfile())
            .ToPowerPointPresentationResult()
            .RequireValue();

        PowerPointTextBox actual = Assert.Single(Assert.Single(imported.Slides).TextBoxes);
        Assert.Equal(2, actual.Paragraphs.Count);
        Assert.Equal("First", actual.Paragraphs[0].Text);
        Assert.Equal("Second", actual.Paragraphs[1].Text);
        Assert.True(actual.Paragraphs[0].Runs[0].Bold);
        Assert.True(actual.Paragraphs[1].Runs[0].Italic);
    }

    [Fact]
    public void LegacyPowerPointSemanticHeadingIsNotPairedWithATextBox() {
        const string html = """
            <section class="officeimo-slide">
              <h2>Slide 1</h2>
              <p>Actual first box</p>
              <p>Actual second box</p>
            </section>
            """;

        using PowerPointPresentation imported = HtmlConversionDocument.Parse(html)
            .ToPowerPointPresentationResult()
            .RequireValue();

        string[] texts = Assert.Single(imported.Slides).TextBoxes.Select(box => box.Text).ToArray();
        Assert.Equal(new[] { "Actual first box", "Actual second box" }, texts);
    }

    [Fact]
    public void OneNoteSemanticHtmlRendersInlineMathFromTheExpression() {
        OneNoteSection section = CreateOneNoteSection(out OneNoteParagraph paragraph);
        paragraph.AddMath(OfficeMath.Fraction(OfficeMath.Identifier("x"), OfficeMath.Number("2")));

        string html = section.ToHtmlDocument();

        Assert.Contains("class=\"officeimo-onenote-math\"", html, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-math-format=\"latex\"", html, StringComparison.Ordinal);
        Assert.Contains("\\frac{x}{2}", html, StringComparison.Ordinal);
    }

    [Fact]
    public void OneNoteSemanticHtmlDoesNotResolveMissingBinaryPayloads() {
        OneNoteSection section = CreateOneNoteSection(out _);
        OneNotePage page = Assert.Single(section.Pages);
        page.DirectContent.Add(new OneNoteImage { FileName = "missing.png", AltText = "Missing" });
        int resolverCalls = 0;

        string html = section.ToHtmlDocument(new OfficeIMO.OneNote.Markdown.OneNoteMarkdownOptions {
            AssetUriResolver = element => {
                resolverCalls++;
                return element.Payload!.Length.ToString();
            }
        });

        Assert.Equal(0, resolverCalls);
        Assert.Contains("officeimo-onenote-image-placeholder", html, StringComparison.Ordinal);
    }

    [Fact]
    public void OneNoteHtmlDiagnosticsSuppressPreservedStylesButRetainSpacingLoss() {
        OneNoteSection section = CreateOneNoteSection(out OneNoteParagraph paragraph);
        paragraph.Runs.Add(new OneNoteTextRun { Text = "Styled" });
        paragraph.Runs[0].Style.FontFamily = "Aptos";
        paragraph.Runs[0].Style.Underline = true;

        HtmlTextConversionResult styled = section.ToHtmlDocumentResult();
        Assert.DoesNotContain(styled.Report.Diagnostics,
            diagnostic => diagnostic.Code is "ONENOTE_MARKDOWN_FORMATTING_SIMPLIFIED" or "ONENOTE_HTML_FORMATTING_SIMPLIFIED");

        paragraph.Style.SpaceBefore = 12D;
        HtmlTextConversionResult spaced = section.ToHtmlDocumentResult();
        Assert.Contains(spaced.Report.Diagnostics,
            diagnostic => diagnostic.Code == "ONENOTE_HTML_FORMATTING_SIMPLIFIED"
                && diagnostic.LossKind == OfficeConversionLossKind.Approximation);
    }

    private static OneNoteSection CreateOneNoteSection(out OneNoteParagraph paragraph) {
        var section = new OneNoteSection { Name = "Review" };
        var page = new OneNotePage { Title = "Review" };
        paragraph = new OneNoteParagraph();
        page.DirectContent.Add(paragraph);
        section.Pages.Add(page);
        return section;
    }
}
