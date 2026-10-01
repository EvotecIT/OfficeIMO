using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("avif-opaque")]
    [InlineData("avif-alpha")]
    public void HtmlRender_AvifPreservesIndependentPixelsInScreenAndPdf(string name) {
        string root = Path.Combine(AppContext.BaseDirectory, "Documents", "Html", "Qualification", "StaticPdfGaps");
        byte[] bytes = File.ReadAllBytes(Path.Combine(root, name + ".avif"));
        byte[] reference = File.ReadAllBytes(Path.Combine(root, name + ".rgba"));
        var document = HtmlConversionDocument.Parse("<img width='49' height='33' src='data:image/avif;base64,"
            + Convert.ToBase64String(bytes) + "'>");
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Continuous, ViewportWidth = 49, ViewportHeight = 33,
            Margins = HtmlRenderMargins.All(0), Scale = 1,
            ResourceUrlPolicy = HtmlUrlPolicy.CreateEmbeddedResourceProfile()
        };
        var rendered = HtmlRenderTestDriver.Render(document, options);
        Assert.Empty(rendered.Diagnostics);
        var visual = Assert.Single(EnumerateRenderVisuals(rendered.Pages[0].Visuals).OfType<HtmlRenderImage>());
        Assert.True(OfficeRasterImageDecoder.TryDecode(visual.Bytes, out var raster));
        AssertAvifPixels(reference, raster!);
        var result = document.RenderToPdfResult(HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged,
            HtmlRenderEncoder.Pdf, new HtmlToPdfOptions(options)));
        Assert.Empty(result.Output.Warnings);
        var embedded = Assert.Single(OfficeIMO.Pdf.PdfImageExtractor.ExtractImages(result.ToBytes()));
        Assert.True(OfficeRasterImageDecoder.TryDecode(embedded.Bytes, out var extracted));
        AssertAvifPixels(reference, extracted!);
    }

    private static void AssertAvifPixels(byte[] reference, OfficeRasterImage image) {
        Assert.Equal(49, image.Width); Assert.Equal(33, image.Height);
        Assert.Equal(reference.Length, image.PixelBuffer.Length);
        for (int i = 0; i < reference.Length; i++)
            Assert.InRange(Math.Abs(reference[i] - image.PixelBuffer[i]), 0, i % 4 == 3 ? 0 : 3);
    }
}
