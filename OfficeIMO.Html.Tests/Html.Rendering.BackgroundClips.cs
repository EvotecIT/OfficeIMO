using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Tests.Pdf;
using Xunit;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("linear-gradient(red,red)")]
    [InlineData("radial-gradient(red,red)")]
    [InlineData("conic-gradient(red,red)")]
    public void SquareBackgroundClipConstrainsEveryGradientBackend(string gradient) {
        string html = BackgroundClipFixture(gradient);
        HtmlRenderOptions options = BackgroundClipOptions();
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);
        HtmlRenderPathClipGroup clip = Assert.Single(EnumerateRenderVisuals(rendered.Pages[0].Scene).OfType<HtmlRenderPathClipGroup>(),
            item => item.Source == "div#target:background-gradient:clip");
        Assert.Equal((12D, 12D, 80D, 60D), (clip.X, clip.Y, clip.Width, clip.Height));
        Assert.Equal(OfficeClipPathKind.Rectangle, clip.ClipPath.Kind);

        AssertBackgroundClipPixels(OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing()));
        Assert.True(OfficePngReader.TryDecode(HtmlConversionDocument.Parse(html).ToPng(options), out OfficeRasterImage? png));
        AssertBackgroundClipPixels(png!);
        Assert.Contains("<clipPath", HtmlConversionDocument.Parse(html).ToSvg(options), StringComparison.Ordinal);
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdfBytes(new HtmlToPdfOptions(options));
        AssertBackgroundClipPixels(OfficeDrawingRasterRenderer.Render(PdfCore.PdfReadDocument.Open(pdf).Pages[0].ToDrawing(), scale: 4D / 3D));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SquareBackgroundClipPreservesCroppedRasterAndVectorImages(bool vector) {
        string image = vector
            ? "data:image/svg+xml;base64," + Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(
                "<svg xmlns='http://www.w3.org/2000/svg' width='2' height='2'><rect width='2' height='2' fill='red'/></svg>"))
            : "data:image/png;base64," + Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(255, 0, 0));
        string html = BackgroundClipFixture("url(\"" + image + "\")");
        HtmlRenderOptions options = BackgroundClipOptions();
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);
        AssertBackgroundClipPixels(OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing()));
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdfBytes(new HtmlToPdfOptions(options));
        AssertBackgroundClipPixels(OfficeDrawingRasterRenderer.Render(PdfCore.PdfReadDocument.Open(pdf).Pages[0].ToDrawing(), scale: 4D / 3D));
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.BackgroundImageValueUnsupported);
    }

    [Theory]
    [InlineData(0, 20)]
    [InlineData(80, 0)]
    public void EmptyContentBackgroundClipOmitsEmptyPaintAndPreservesTheOtherLayer(int width, int height) {
        foreach (bool paddingColor in new[] { true, false }) {
            string html = "<style>html,body{margin:0}</style><div id='target' style='width:" + width + "px;height:" + height + "px;padding:8px;border:4px solid black;"
                + "background-image:linear-gradient(red,red),none;background-origin:padding-box;background-clip:content-box,"
                + (paddingColor ? "padding-box" : "content-box") + ";background-color:lime'></div>";
            HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, BackgroundClipOptions());
            Assert.DoesNotContain(EnumerateRenderVisuals(rendered.Pages[0].Scene).OfType<HtmlRenderShape>(),
                shape => shape.Shape.FillGradient != null || !paddingColor && shape.Shape.FillColor == OfficeColor.Lime);
            Assert.Equal(paddingColor ? OfficeColor.Lime : OfficeColor.White,
                OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing()).GetPixel(8, 8));
        }
    }

    [Theory]
    [InlineData("none", "content-box,padding-box", false)]
    [InlineData("none,none", "padding-box,content-box,border-box", false)]
    [InlineData("none,none,none", "content-box,padding-box", false)]
    [InlineData("none,none", "padding-box", true)]
    public void BackgroundColorUsesTheFinalImageLayersRecycledClip(string images, string clips, bool paddingColor) {
        string html = "<style>html,body{margin:0}</style><div id='target' style='width:80px;height:60px;padding:8px;border:4px solid black;"
            + "background-image:" + images + ";background-clip:" + clips + ";background-color:red'></div>";
        HtmlRenderOptions options = BackgroundClipOptions();
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);
        HtmlRenderShape fill = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(),
            shape => shape.Source == "div#target" && shape.Shape.FillColor == OfficeColor.Red);
        Assert.Equal(paddingColor ? (4D, 4D, 96D, 76D) : (12D, 12D, 80D, 60D),
            (fill.X, fill.Y, fill.Width, fill.Height));
        OfficeColor expected = paddingColor ? OfficeColor.Red : OfficeColor.White;
        Assert.Equal(expected, OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing()).GetPixel(8, 20));
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdfBytes(new HtmlToPdfOptions(options));
        Assert.Equal(expected, OfficeDrawingRasterRenderer.Render(PdfCore.PdfReadDocument.Open(pdf).Pages[0].ToDrawing(), scale: 4D / 3D, background: OfficeColor.White).GetPixel(8, 20));
    }

    private static string BackgroundClipFixture(string image) =>
        "<style>html,body{margin:0}</style><div id='target' style='width:80px;height:60px;padding:8px;border:4px solid black;"
        + "background-image:" + image + ",none;background-origin:padding-box;background-clip:content-box,padding-box;"
        + "background-size:100% 100%;background-repeat:no-repeat;background-color:lime'></div>";

    private static HtmlRenderOptions BackgroundClipOptions() => new HtmlRenderOptions {
        Mode = HtmlRenderMode.Continuous,
        ViewportWidth = 120D,
        ViewportHeight = 100D,
        Margins = HtmlRenderMargins.All(0D),
        BackgroundColor = OfficeColor.White,
        ConicGradientQualitySegments = 16
    };

    private static void AssertBackgroundClipPixels(OfficeRasterImage raster) {
        Assert.Equal(OfficeColor.Black, raster.GetPixel(1, 20));
        Assert.Equal(OfficeColor.Lime, raster.GetPixel(8, 20));
        Assert.Equal(OfficeColor.Red, raster.GetPixel(20, 16));
        Assert.Equal(OfficeColor.Lime, raster.GetPixel(95, 20));
    }
}
