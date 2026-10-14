using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    private const string ZeroFontFootnotePixel = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNgYAAAAAMAASsJTYQAAAAASUVORK5CYII=";

    [Theory]
    [InlineData("", false)]
    [InlineData("font-size:16px", false)]
    [InlineData("", true)]
    public void HtmlFootnote_ZeroFontParentSuppressesDefaultCallAndMarkerWithoutLosingDestinations(string noteStyle, bool columnNotes) {
        string html = "<style>body{margin:0;font:16px/20px FootnoteFixture}p{margin:0}.note{float:footnote}</style>"
            + (columnNotes ? "<style>.columns{column-count:2;column-gap:20px;height:100px}.note{float-reference:column}</style>" : string.Empty)
            + "<section class='columns'><p style='font-size:0'>HIDDEN<span id='zero-note' class='note' style='" + noteStyle
            + "'>ZeroNote</span></p><p>Visible<span id='positive-note' class='note'>PositiveNote</span></p></section>";
        HtmlConversionDocument document = HtmlConversionDocument.Parse(html);
        HtmlRenderOptions options = ZeroFontFootnoteOptions();
        HtmlRenderVisual[] visuals = ZeroFontFootnoteVisuals(HtmlRenderTestDriver.Render(document, options));
        HtmlRenderText[] text = visuals.OfType<HtmlRenderText>().ToArray();

        Assert.DoesNotContain(text, item => item.Source == "span#zero-note:footnote-call");
        Assert.Equal(noteStyle.Length > 0, text.Any(item => item.Text == "ZeroNote"));
        Assert.Equal(noteStyle.Length > 0, text.Any(item => item.Source == "span#zero-note:footnote-marker"));
        Assert.Contains(text, item => item.Text == "Visible");
        Assert.Contains(text, item => item.Text == "PositiveNote");
        Assert.Contains(text, item => item.Source == "span#positive-note:footnote-call" && item.Font.Size == 11.2D);
        Assert.All(text, item => Assert.True(item.Font.Size > 0D));
        AssertFootnoteDestinations(visuals, 1);
        AssertFootnoteDestinations(visuals, 2);

        string pdfText = OfficeIMO.Pdf.PdfReadDocument.Open(document.ToPdfBytes(new HtmlToPdfOptions(options))).ExtractText();
        Assert.DoesNotContain("HIDDEN", pdfText, StringComparison.Ordinal);
        Assert.Contains("PositiveNote", pdfText, StringComparison.Ordinal);
        Assert.Equal(noteStyle.Length > 0, pdfText.Contains("ZeroNote", StringComparison.Ordinal));
    }

    [Fact]
    public void HtmlFootnote_ZeroFontParentPreservesAuthoredPositiveCallMarkerAndDescendantFonts() {
        const string html = "<style>body{margin:0;font:16px/20px FootnoteFixture}.note{float:footnote}"
            + ".note::footnote-call{content:'C';font-size:12px}.note::footnote-marker{content:'M';font-size:10px}</style>"
            + "<p style='font-size:0'>HIDDEN<span class='note'>HIDDEN_NOTE<span style='font-size:14px'>Restored</span></span></p>";
        HtmlConversionDocument document = HtmlConversionDocument.Parse(html);
        HtmlRenderOptions options = ZeroFontFootnoteOptions();
        HtmlRenderVisual[] visuals = ZeroFontFootnoteVisuals(HtmlRenderTestDriver.Render(document, options));
        HtmlRenderText[] text = visuals.OfType<HtmlRenderText>().ToArray();

        Assert.Equal(12D, Assert.Single(text, item => item.Text == "C").Font.Size);
        Assert.Equal(10D, Assert.Single(text, item => item.Text == "M").Font.Size);
        Assert.Equal(14D, Assert.Single(text, item => item.Text == "Restored").Font.Size);
        Assert.DoesNotContain(text, item => item.Text.Contains("HIDDEN", StringComparison.Ordinal));
        AssertFootnoteDestinations(visuals, 1);
        string pdfText = OfficeIMO.Pdf.PdfReadDocument.Open(document.ToPdfBytes(new HtmlToPdfOptions(options))).ExtractText();
        Assert.Contains("Restored", pdfText, StringComparison.Ordinal);
        Assert.DoesNotContain("HIDDEN", pdfText, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void HtmlFootnote_ZeroFontMarkersSuppressDefaultTextAndContinuationsButPreserveImages(
        bool authoredMarker, bool imageMarker) {
        string markerContent = authoredMarker ? ";content:'HIDDEN_MARKER'"
            + (imageMarker ? " url('data:image/png;base64," + ZeroFontFootnotePixel + "')" : string.Empty) : string.Empty;
        string html = "<style>@page{size:240px 160px;margin:10px}body,p{margin:0;font:12px/16px FootnoteFixture}"
            + ".note{float:footnote;display:inline-block;font-size:10px;line-height:12px}"
            + ".note::footnote-marker{font-size:0" + markerContent + "}</style><p>Body<span class='note'>"
            + string.Concat(Enumerable.Range(1, 12).Select(index => "Note" + index.ToString("D2") + "<br>"))
            + "</span></p>";
        HtmlConversionDocument document = HtmlConversionDocument.Parse(html);
        HtmlRenderOptions options = ZeroFontFootnoteOptions();
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(document, options);
        HtmlRenderVisual[] visuals = ZeroFontFootnoteVisuals(rendered);
        HtmlRenderText[] text = visuals.OfType<HtmlRenderText>().ToArray();

        Assert.True(rendered.Pages.Count >= 2);
        Assert.All(text, item => Assert.True(item.Font.Size > 0D));
        Assert.DoesNotContain(text, item => item.Source?.Contains("footnote-marker", StringComparison.Ordinal) == true);
        Assert.Contains(text, item => item.Text == "Note01");
        Assert.Contains(text, item => item.Text == "Note12");
        if (imageMarker) Assert.True(visuals.OfType<HtmlRenderImage>().Count() >= 2);
        else Assert.Empty(visuals.OfType<HtmlRenderImage>());
        AssertFootnoteDestinations(visuals, 1);
        string pdfText = OfficeIMO.Pdf.PdfReadDocument.Open(document.ToPdfBytes(new HtmlToPdfOptions(options))).ExtractText();
        Assert.Contains("Note12", pdfText, StringComparison.Ordinal);
        Assert.DoesNotContain("HIDDEN_MARKER", pdfText, StringComparison.Ordinal);
        Assert.DoesNotContain("cont.", pdfText, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlFootnote_ZeroFontGeneratedCallKeepsItsImageAndPositiveNoteDescendant() {
        string html = "<style>body{margin:0;font:16px/20px FootnoteFixture}.note{float:footnote}"
            + ".note::footnote-call{font-size:0;content:'HIDDEN_CALL' url('data:image/png;base64," + ZeroFontFootnotePixel
            + "')}</style><p style='font-size:0'>HIDDEN<span class='note'><span style='font-size:14px'>Restored</span></span></p>";
        HtmlConversionDocument document = HtmlConversionDocument.Parse(html);
        HtmlRenderOptions options = ZeroFontFootnoteOptions();
        HtmlRenderVisual[] visuals = ZeroFontFootnoteVisuals(HtmlRenderTestDriver.Render(document, options));
        HtmlRenderText[] text = visuals.OfType<HtmlRenderText>().ToArray();

        Assert.Single(visuals.OfType<HtmlRenderImage>());
        Assert.Equal(14D, Assert.Single(text, item => item.Text == "Restored").Font.Size);
        Assert.DoesNotContain(text, item => item.Text.Contains("HIDDEN", StringComparison.Ordinal));
        Assert.All(text, item => Assert.True(item.Font.Size > 0D));
        AssertFootnoteDestinations(visuals, 1);
        string pdfText = OfficeIMO.Pdf.PdfReadDocument.Open(document.ToPdfBytes(new HtmlToPdfOptions(options))).ExtractText();
        Assert.Contains("Restored", pdfText, StringComparison.Ordinal);
        Assert.DoesNotContain("HIDDEN", pdfText, StringComparison.Ordinal);
    }

    private static HtmlRenderOptions ZeroFontFootnoteOptions() {
        var options = new HtmlRenderOptions { Mode = HtmlRenderMode.Paged, HonorCssPageRules = true, Margins = HtmlRenderMargins.All(0D),
            ResourceUrlPolicy = HtmlUrlPolicy.CreateEmbeddedResourceProfile() };
        options.Fonts.Add("FootnoteFixture", ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(
            "HIDDEN_CALLMARKERZeroNotePositiveVisibleRestoredBody0123456789()cont. ".Distinct().Select(value => (int)value).ToArray()));
        return options;
    }

    private static HtmlRenderVisual[] ZeroFontFootnoteVisuals(HtmlRenderDocument document) =>
        document.Pages.SelectMany(page => EnumerateRenderVisuals(page.Scene)).ToArray();

    private static void AssertFootnoteDestinations(IEnumerable<HtmlRenderVisual> visuals, int number) {
        HtmlRenderNamedDestination[] destinations = visuals.OfType<HtmlRenderNamedDestination>().ToArray();
        Assert.Single(destinations, destination => destination.Name == "officeimo-footnote-call-" + number);
        Assert.Single(destinations, destination => destination.Name == "officeimo-footnote-note-" + number);
    }
}
