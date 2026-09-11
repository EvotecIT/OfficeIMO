using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using System.Threading;
using Xunit;

namespace OfficeIMO.Tests;

public class HtmlPdfTextPlacement {
    [Fact]
    public void LayoutDefersUncoveredEmbeddedGlyphsUntilAutomaticFallbackSelection() {
        byte[] primaryBytes = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fonts", "SourceSerif4-Regular.otf"));
        var options = new HtmlToPdfOptions {
            PdfOptions = new OfficeIMO.Pdf.PdfOptions().EmbedStandardFont(
                OfficeIMO.Pdf.PdfStandardFont.Helvetica, primaryBytes, "Primary")
        };
        const string expected = "office\u012C continued";
        byte[] bytes = HtmlConversionDocument.Parse("<p>" + expected + "</p>").ToPdfBytes(options);
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        Assert.Contains(expected, pdf.GetPage(1).Text, StringComparison.Ordinal);
    }

    [Fact]
    public void PrecomputedSceneKeepsTextAndLinkWithinItsAllocatedAdvance() {
        const string uri = "https://example.test/details";
        var rendered = HtmlRenderTestDriver.Render("<p style='font:24px Arial'><a href='" + uri + "'>NORTHWIND</a> follows</p>");
        var result = HtmlPdfRenderedConverter.CreatePdf(rendered, new HtmlToPdfOptions(), CancellationToken.None);
        byte[] bytes = result.Document.ToBytes();
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        var letters = pdf.GetPage(1).Letters.Where(letter => !string.IsNullOrWhiteSpace(letter.Value)).ToArray();
        Assert.Equal("NORTHWINDfollows", string.Concat(letters.Select(letter => letter.Value)));
        Assert.True(letters[9].StartBaseLine.X >= letters[8].EndBaseLine.X);
        var link = Assert.Single(OfficeIMO.Pdf.PdfDocumentReadResult.Load(bytes).GetLinksByUri(uri));
        Assert.InRange(link.Width, letters[8].EndBaseLine.X - letters[0].StartBaseLine.X - 0.1D,
            letters[8].EndBaseLine.X - letters[0].StartBaseLine.X + 0.1D);
    }

    [Fact]
    public void LayoutUsesConfiguredEmbeddedFallbacksBeforeMeasuringTheirGlyphs() {
        byte[] primaryBytes = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fonts", "SourceSerif4-Regular.otf"));
        byte[] fallbackBytes = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fonts", "RobotoFlex.ttf"));
        var primary = OfficeIMO.Pdf.PdfOpenTypeCffFontProgram.Parse(primaryBytes, "Primary");
        var fallback = OfficeIMO.Pdf.PdfTrueTypeFontProgram.Parse(fallbackBytes, "Fallback");
        int scalar = Enumerable.Range(0x20, 0x3000 - 0x20).First(value =>
            !char.IsControl((char)value) && !char.IsWhiteSpace((char)value)
            && (!primary.TryGetGlyphId(value, out int primaryGlyph) || primaryGlyph == 0)
            && fallback.TryGetGlyphId(value, out int fallbackGlyph) && fallbackGlyph > 0);
        string text = "office" + char.ConvertFromUtf32(scalar);
        var options = new OfficeIMO.Pdf.PdfOptions()
            .EmbedStandardFont(OfficeIMO.Pdf.PdfStandardFont.Helvetica, primaryBytes, "Primary")
            .RegisterEmbeddedFontFallbacks(new OfficeIMO.Pdf.PdfEmbeddedFontFallbackSet(
                new[] { new OfficeIMO.Pdf.PdfEmbeddedFontFallbackCandidate("Fallback", fallbackBytes) },
                new[] { OfficeIMO.Pdf.PdfStandardFont.TimesRoman }));
        byte[] bytes = HtmlConversionDocument.Parse("<p>" + text + "</p>").ToPdfBytes(new HtmlToPdfOptions {
            PdfOptions = options,
            TextFallbacks = OfficeIMO.Pdf.PdfTextFallbackFeatures.None
        });
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        Assert.Contains(text, pdf.GetPage(1).Text, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("Arial")]
    [InlineData("Times New Roman")]
    [InlineData("Courier New")]
    public void PositionedCharactersUseTheWrittenFontsAdvance(string family) {
        byte[] bytes = HtmlConversionDocument.Parse(
            $"<p style='font-family:{family};font-size:24px;font-weight:bold;letter-spacing:1px'>NORTHWIND</p>")
            .ToPdfBytes();
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        var letters = pdf.GetPage(1).Letters.Where(letter => !string.IsNullOrWhiteSpace(letter.Value)).ToArray();
        Assert.Equal("NORTHWIND", string.Concat(letters.Select(letter => letter.Value)));
        for (int index = 1; index < letters.Length; index++) {
            double gap = letters[index].StartBaseLine.X - letters[index - 1].EndBaseLine.X;
            Assert.InRange(gap, 0.70D, 0.80D);
        }
    }

    [Theory]
    [InlineData("<strong>46.75 hours recorded.</strong> Three entries are approved.", "46.75 hours recorded. Three entries are approved.")]
    [InlineData("Confirm reference <strong>00130</strong>.", "Confirm reference 00130.")]
    [InlineData("<span>Leading</span> <em>and</em> <strong>trailing</strong> spaces.", "Leading and trailing spaces.")]
    public void InlineBoundariesPreserveSpacesAndDoNotOverlap(string body, string expected) {
        byte[] bytes = HtmlConversionDocument.Parse($"<p style='font:14px Arial'>{body}</p>").ToPdfBytes();
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        var letters = pdf.GetPage(1).Letters.ToArray();
        Assert.Equal(expected, string.Concat(letters.Select(letter => letter.Value)));
        for (int index = 1; index < letters.Length; index++) {
            Assert.True(letters[index].StartBaseLine.X >= letters[index - 1].EndBaseLine.X - 0.02D,
                $"'{letters[index - 1].Value}' overlaps '{letters[index].Value}'.");
        }
    }

    [Fact]
    public void PositionedTextRetainsStandardFontPunctuation() {
        byte[] bytes = HtmlConversionDocument.Parse("<p>September · Approved • € £</p>").ToPdfBytes();
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        Assert.Contains("September · Approved • € £", pdf.GetPage(1).Text, StringComparison.Ordinal);
    }
}
