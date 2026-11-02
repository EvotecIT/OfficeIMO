using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.TestAssets;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("Scoped, Literal")]
    [InlineData("Scoped; Literal")]
    [InlineData("Scoped\\Literal\"Face")]
    public void HtmlFontFallback_ResolvedLiteralFamiliesKeepTheirMetricsAndPdfFont(string family) {
        byte[] bytes = ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs('0', 'B');
        OfficeTrueTypeFont font = OfficeTrueTypeFont.TryLoad(bytes)!;
        string request = "\"" + family.Replace("\\", "\\\\").Replace("\"", "\\\"") + "\"";
        string html = "<style>body,p{margin:0}body{font-family:" + request
            + ";font-size:20px;line-height:24px}</style><p>BB</p><div style='margin-left:3ch'>B</div>";
        var options = new HtmlToPdfOptions {
            Mode = HtmlRenderMode.Continuous,
            AllowSystemFontFallback = false,
            FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss,
            TextFallbacks = PdfCore.PdfTextFallbackFeatures.None,
            ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic()
        };
        options.Fonts.Add(family, bytes);
        var document = HtmlConversionDocument.Parse(html);
        var text = HtmlRenderTestDriver.Render(document, options).Pages[0].Visuals.OfType<HtmlRenderText>().ToArray();
        HtmlRenderText first = Assert.Single(text, item => item.Text == "BB");
        HtmlRenderText shifted = Assert.Single(text, item => item.Text == "B");

        Assert.Equal(font.Measure("BB", 20D), first.TextAdvanceWidth!.Value, 6);
        Assert.Equal(font.Measure("0", 20D) * 3D, shifted.X - first.X, 6);
        byte[] pdf = (document.ToPdf(options)).ToBytes();
        Assert.Contains("BB", PdfCore.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
        Assert.Equal(1, PdfCore.PdfFontInspector.Inspect(PdfCore.PdfReadDocument.Open(pdf)).EmbeddedFontCount);
    }
}
