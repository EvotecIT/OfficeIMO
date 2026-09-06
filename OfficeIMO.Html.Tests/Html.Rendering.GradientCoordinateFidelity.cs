using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData(40, 40, 222, 33)]
    [InlineData(340, 40, 74, 181)]
    [InlineData(40, 150, 184, 71)]
    [InlineData(340, 150, 36, 219)]
    [InlineData(192, 96, 128, 127)]
    public void HtmlLinearGradient_NonSquareColorFieldMatchesIndependentPrintSamples(int x, int y, int red, int blue) {
        // Reference: 384x192 CSS px, Chromium print PDF independently rasterized
        // at 96 DPI. Interior samples allow one level for pixel-center rounding.
        const string html = "<style>@page{size:384px 192px;margin:0}html,body{margin:0}</style>"
            + "<div style='width:384px;height:192px;background:linear-gradient(125deg,#ff0000,#0000ff)'></div>";
        var options = new HtmlRenderOptions { Mode = HtmlRenderMode.Paged, Margins = HtmlRenderMargins.All(0D) };
        HtmlConversionDocument document = HtmlConversionDocument.Parse(html);
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(document, options);
        OfficeRasterImage direct = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        Assert.Empty(rendered.Diagnostics);
        AssertGradientSample(direct.GetPixel(x, y), red, blue);

        string svg = document.ToSvg(options);
        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? imported, out int unsupported));
        Assert.Equal(0, unsupported);
        AssertGradientSample(OfficeDrawingRasterRenderer.Render(imported!).GetPixel(x, y), red, blue);

        byte[] pdf = document.ToPdfBytes(new HtmlToPdfOptions(options));
        OfficeDrawing pdfDrawing = OfficeIMO.Pdf.PdfPageImageRenderer.RenderPage(pdf);
        OfficeRasterImage pdfImage = OfficeDrawingRasterRenderer.Render(pdfDrawing,
            new OfficeDrawingRasterRenderOptions { Scale = 96D / 72D });
        AssertGradientSample(pdfImage.GetPixel(x, y), red, blue);
    }

    [Fact]
    public void HtmlLinearGradient_CornerLengthStopsUseThePerpendicularPhysicalLine() {
        const string html = "<div style='width:200px;height:100px;background:linear-gradient(to bottom right,red 0px,blue 89.4427191px,blue)'></div>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            ViewportWidth = 200D, Margins = HtmlRenderMargins.All(0D)
        });
        OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        // The midpoint of the magic-corner gradient is half its physical line,
        // 89.4427191px, rather than half the rectangle diagonal (111.8px).
        Assert.True(image.GetPixel(100, 50).B >= 254);
        Assert.InRange(image.GetPixel(50, 25).R, 124, 128);
    }

    private static void AssertGradientSample(OfficeColor actual, int red, int blue) {
        Assert.InRange((int)actual.R, red - 1, red + 1);
        Assert.Equal(0, actual.G);
        Assert.InRange((int)actual.B, blue - 1, blue + 1);
        Assert.Equal(255, actual.A);
    }
}
