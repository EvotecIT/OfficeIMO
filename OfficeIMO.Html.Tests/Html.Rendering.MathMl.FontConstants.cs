using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlMathMlUsesScopedMathConstantsAndRetainsOriginalPdfSemantics() {
        const string html = "<body style='margin:0;font:20px ScopedMath'><math style='font-family:inherit'>"
            + "<mfrac><mtext>x</mtext><mn>2</mn></mfrac></math></body>";
        byte[] font = ManagedTextShapingTestAssets.CreateMathFont();
        var renderOptions = new HtmlRenderOptions { ViewportWidth = 160, ViewportHeight = 160,
            Margins = HtmlRenderMargins.All(0), AllowSystemFontFallback = false };
        renderOptions.Fonts.Add("ScopedMath", font);
        var rendered = HtmlRenderTestDriver.Render(html, renderOptions);
        var math = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderDrawing>());
        Assert.All(math.Drawing.Elements.OfType<OfficeDrawingText>(), t => Assert.Equal(14D, t.Font.Size, 6));
        Assert.Equal(1.36D, Assert.Single(math.Drawing.Elements.OfType<OfficeDrawingShape>()).Shape.StrokeWidth, 6);
        var pdfOptions = new HtmlToPdfOptions { PageSize = new OfficePageSize(2, 2), HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0), AllowSystemFontFallback = false };
        pdfOptions.Fonts.Add("ScopedMath", font);
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdfBytes(pdfOptions);
        Assert.Contains("(x)/(2)", OfficeIMO.Pdf.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
        Assert.Empty(OfficeIMO.Pdf.PdfImageExtractor.ExtractImages(pdf));
        OfficeDrawing reopened = OfficeIMO.Pdf.PdfPageImageRenderer.RenderPage(pdf);
        var raster = OfficeDrawingRasterRenderer.Render(reopened, scale: 3);
        Assert.True(raster.Width > 0 && raster.Height > 0);
    }
}
