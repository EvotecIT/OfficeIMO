using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.TestAssets;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlPdfSvgFontTests {
    private const string PolishText = "Zażółć gęślą jaźń";

    [Theory]
    [InlineData("viewBox='0 0 400 80'", "", "")]
    [InlineData("", "<svg width='400' height='80'>", "</svg>")]
    [InlineData("", "<g opacity='0.75'>", "</g>")]
    public void SvgTextRegistersExplicitUnicodeFontsInsideDrawingGroups(string rootAttributes, string beforeText, string afterText) {
        byte[] font = ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(PolishText.Select(character => (int)character).Distinct().ToArray());
        var options = new HtmlToPdfOptions {
            TextFallbacks = PdfCore.PdfTextFallbackFeatures.None,
            ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic()
        };
        options.Fonts.Add("Chart Unicode", font);
        string html = "<svg xmlns='http://www.w3.org/2000/svg' width='400' height='80' " + rootAttributes + ">"
            + beforeText + "<text x='8' y='30' font-family='Chart Unicode' font-size='20'>" + PolishText + "</text>" + afterText + "</svg>";

        byte[] bytes = HtmlConversionDocument.Parse(html).ToPdfBytes(options);

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        Assert.Contains(PolishText, pdf.GetPage(1).Text, StringComparison.Ordinal);
        Assert.Contains(PolishText, PdfCore.PdfReadDocument.Open(bytes).ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void SvgUnicodeTextStillRequiresCoverageWhenHostFontsAreDisabled() {
        var options = new HtmlToPdfOptions {
            ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic()
        };
        const string html = "<svg xmlns='http://www.w3.org/2000/svg' width='200' height='60'><svg width='200' height='60'><text x='5' y='30'>ęłź</text></svg></svg>";

        Assert.ThrowsAny<ArgumentException>(() => HtmlConversionDocument.Parse(html).ToPdfBytes(options));
    }
}
