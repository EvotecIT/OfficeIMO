using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    public void HtmlPdf_LibraryFontEmbeddingUsesTheSameMetricsForLayoutAndPaint(bool portable, bool optOut) {
        if (!portable && !optOut && !new[] { "Arial", "Liberation Sans", "Arimo", "Helvetica Neue", "Aptos" }
                .Any(name => PdfCore.PdfEmbeddedFontFamily.TryFromSystem(name, out _))) return;
        var options = new HtmlToPdfOptions {
            Margins = HtmlRenderMargins.All(0),
            ResourcePolicy = portable ? PdfCore.PdfResourcePolicy.CreatePortableDeterministic()
                : PdfCore.PdfResourcePolicy.CreateLibraryFontDefault()
        };
        if (optOut) options.PdfOptions.EmbedDefaultFonts = false;
        HtmlPdfRenderResult result = HtmlPdfRenderedConverter.Convert(
            HtmlConversionDocument.Parse("<p style='margin:0;font:16px Arial'>WWWiii</p>"), options);
        using PdfCore.PdfDocument model = result.Document;
        byte[] bytes = model.ToBytes();
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        var letters = pdf.GetPage(1).Letters;
        HtmlRenderText text = Assert.Single(result.RenderResult!.Document.Pages[0].Visuals.OfType<HtmlRenderText>());

        Assert.Equal("WWWiii", string.Concat(letters.Select(letter => letter.Value)));
        Assert.Equal((text.TextPaintWidth ?? text.TextAdvanceWidth ?? text.Width) * .75D,
            letters.Last().EndBaseLine.X - letters.First().StartBaseLine.X, 2);
        Assert.Equal(!portable && !optOut, model.Options.EmbeddedFonts.Count > 0);
        if (portable) Assert.Equal(0, PdfCore.PdfDiagnostics.Analyze(bytes).EmbeddedFontCount);
        else if (!optOut) Assert.True(PdfCore.PdfDiagnostics.Analyze(bytes).EmbeddedFontCount > 0);
    }
}
