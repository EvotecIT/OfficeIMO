using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.TestAssets;
using System.Text;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlPdf_InstalledFontMeasurementKeepsPrintedFlexLinksTogether() {
        string? installedFamily = new[] { "Trebuchet MS", "Arial", "Calibri", "Liberation Sans", "DejaVu Sans" }
            .FirstOrDefault(candidate => PdfCore.PdfEmbeddedFontFamily.TryFromSystem(candidate, out _));
        if (installedFamily == null) return;

        string html = "<style>body{margin:0;font-family:'Missing Document Font','" + installedFamily + "',sans-serif}"
            + "ul{display:flex;flex-wrap:wrap;width:760px;margin:0;padding:0;list-style:none}"
            + "li{display:inline-block;margin-left:11px;padding-right:16px}"
            + "a{display:block;font-size:15px;line-height:1.5}"
            + "@media print{a::after{content:' (' attr(href) ')'}}"
            + "</style><ul><li><a href='https://www.w3.org/WAI/'>Home</a></li>"
            + "<li><a href='https://www.w3.org/WAI/design-develop/'>Design &amp; Develop</a></li></ul>";
        var options = new HtmlToPdfOptions {
            PageSize = new OfficeIMO.Drawing.OfficePageSize(816D / HtmlRenderOptions.CssPixelsPerInch, 900D / HtmlRenderOptions.CssPixelsPerInch),
            Margins = HtmlRenderMargins.All(0D),
            HonorCssPageRules = false
        };
        options.ResourcePolicy.AllowDocumentFontEmbedding = true;

        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdfBytes(options);
        string extracted = PdfCore.PdfReadDocument.Open(pdf).ExtractText();

        Assert.Contains("Home (https://www.w3.org/WAI/)", extracted, StringComparison.Ordinal);
        Assert.Contains("Design & Develop (https://www.w3.org/WAI/design-develop/)", extracted, StringComparison.Ordinal);
        Assert.True(PdfCore.PdfDiagnostics.Analyze(pdf).EmbeddedFontCount > 0);
    }

    [Fact]
    public void HtmlPdf_InstalledFontMeasurementUsesNextCssFamilyForMissingWebGlyph() {
        string? installedFamily = new[] { "Trebuchet MS", "Arial", "Calibri", "Liberation Sans", "DejaVu Sans" }
            .FirstOrDefault(candidate => PdfCore.PdfEmbeddedFontFamily.TryFromSystem(candidate, out _));
        if (installedFamily == null) return;

        string html = "<style>" + CreatePortableEmbeddedFontFaceCss("Scoped Web")
            + "p{font-family:'Scoped Web','" + installedFamily + "';font-size:20px}</style>"
            + "<p>AΩ</p>";
        var options = new HtmlToPdfOptions {
            TextFallbacks = PdfCore.PdfTextFallbackFeatures.None
        };
        options.ResourcePolicy.AllowDocumentFontEmbedding = true;

        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdfBytes(options);
        string rawPdf = Encoding.ASCII.GetString(pdf);

        Assert.Contains("/BaseFont /ScopedWeb-Regular", rawPdf, StringComparison.Ordinal);
        Assert.Contains("AΩ", PdfCore.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
        Assert.True(PdfCore.PdfDiagnostics.Analyze(pdf).EmbeddedFontCount >= 2);
    }

    [Fact]
    public void HtmlPdf_FontMeasurementUsesTheSameCoveringNamedFallbackAsPdfOutput() {
        byte[] latin = ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs('A');
        byte[] greek = ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs('Ω');
        var options = new HtmlToPdfOptions {
            TextFallbacks = PdfCore.PdfTextFallbackFeatures.None,
            Margins = HtmlRenderMargins.All(0D)
        };
        options.ResourcePolicy.AllowDocumentFontEmbedding = true;
        options.PdfOptions.RegisterNamedFontFamily(new PdfCore.PdfEmbeddedFontFamily("LatinOnly", latin));
        options.PdfOptions.RegisterNamedFontFamily(new PdfCore.PdfEmbeddedFontFamily("GreekFallback", greek));

        HtmlPdfRenderResult result = HtmlPdfRenderedConverter.Convert(
            HtmlConversionDocument.Parse("<p style='margin:0;font:40px LatinOnly,GreekFallback'>ΩΩ</p>"), options);
        HtmlRenderText text = Assert.Single(result.RenderResult!.Document.Pages[0].Visuals.OfType<HtmlRenderText>());
        double expected = Assert.IsType<double>(PdfCore.PdfWriter.MeasurePositionedText(
            new PdfCore.PdfTextRun("ΩΩ", fontSize: 40D, fontFamily: "GreekFallback"),
            result.Document.Options));

        Assert.Equal(expected, Assert.IsType<double>(text.TextAdvanceWidth), 3);
        Assert.Contains("ΩΩ", PdfCore.PdfReadDocument.Open(result.Document.ToBytes()).ExtractText(), StringComparison.Ordinal);
    }
}
