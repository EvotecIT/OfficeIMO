using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("50% 50%")]
    [InlineData("0% 0%")]
    public void HtmlConicGradient_VariableAlphaKeepsVectorPdf(string center) {
        string html = "<div style='width:180px;height:180px;background:conic-gradient(at " + center
            + ",rgba(255 0 0 / 25%),rgba(0 0 255 / 75%))'></div>";
        var options = new HtmlRenderOptions { Mode = HtmlRenderMode.Continuous,
            ViewportWidth = 200D, Margins = HtmlRenderMargins.All(0D), BackgroundColor = OfficeColor.Transparent };
        var result = HtmlConversionDocument.Parse(html).ToPdfDocumentResult(new HtmlToPdfOptions(options));
        byte[] pdf = result.ToBytes();

        Assert.DoesNotContain(result.Report.Warnings, warning => warning.Code == "HtmlPdfDrawingEffectRasterized");
        Assert.Empty(PdfCore.PdfImageExtractor.ExtractImages(pdf));
        Assert.Contains("/S /Luminosity", System.Text.Encoding.ASCII.GetString(pdf), StringComparison.Ordinal);
        var drawing = PdfCore.PdfReadDocument.Open(pdf).Pages[0].ToDrawing();
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing);
        Assert.InRange(raster.GetPixel(60, 60).A, (byte)64, (byte)192);
    }
}
