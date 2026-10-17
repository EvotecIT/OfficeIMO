using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Pdf;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    private const string PagedImageHtml = "<style>@page{size:160px 120px;margin:10px;"
        + "@bottom-center{content:'Page ' counter(page) ' of ' counter(pages)}}"
        + "body{margin:0;font:12px Arial}p{margin:0}</style><p>First page</p>"
        + "<p style='break-before:page'>Second page</p>";

    [Theory]
    [InlineData(OfficeImageExportFormat.Png)]
    [InlineData(OfficeImageExportFormat.Jpeg)]
    [InlineData(OfficeImageExportFormat.Svg)]
    public void HtmlImages_DefaultPagesAreIdenticalToPdfPages(OfficeImageExportFormat format) {
        var document = HtmlConversionDocument.Parse(PagedImageHtml);
        var options = new HtmlRenderOptions { AllowSystemFontFallback = false };
        var actual = document.ExportImages(format, options);
        var expected = document.ToPdfDocumentResult(new HtmlToPdfOptions(options)).ExportImages(format,
            new PdfImageExportOptions { Scale = 96D / 72D });

        Assert.Equal(2, actual.Count);
        Assert.Equal(expected.Count, actual.Count);
        for (int index = 0; index < actual.Count; index++) {
            Assert.Equal(160, actual[index].Width);
            Assert.Equal(120, actual[index].Height);
            Assert.Equal(expected[index].Bytes, actual[index].Bytes);
        }
    }

    [Theory]
    [InlineData(96D, 160, 120)]
    [InlineData(144D, 240, 180)]
    public async Task HtmlImages_SelectionAndAsyncStreamingUseTheSamePdfScene(double dpi, int width, int height) {
        var document = HtmlConversionDocument.Parse(PagedImageHtml);
        var options = new HtmlRenderOptions { AllowSystemFontFallback = false, TargetDpi = dpi };
        var all = document.ExportImages(OfficeImageExportFormat.Png, options);
        var pdfPages = document.ToPdfDocumentResult(new HtmlToPdfOptions(options)).ExportImages(
            OfficeImageExportFormat.Png, new PdfImageExportOptions { TargetDpi = dpi });
        Assert.Equal(all.Count, pdfPages.Count);
        for (int index = 0; index < all.Count; index++) Assert.Equal(pdfPages[index].Bytes, all[index].Bytes);
        var streamed = new List<OfficeImageExportResult>();
        await document.ExportImagesAsync(OfficeImageExportFormat.Png, async (image, token) => {
            await Task.Yield();
            token.ThrowIfCancellationRequested();
            streamed.Add(image);
        }, options);
        var retained = HtmlRenderEngine.Execute(document, HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged,
            HtmlRenderEncoder.Png, options).WithPageSet(HtmlRenderPageSet.Page(1)));
        var selected = retained.ExportImage();
        Assert.Equal(width, selected.Width);
        Assert.Equal(height, selected.Height);
        Assert.Equal("Page 2", selected.Name);
        Assert.Equal(all[1].Bytes, selected.Bytes);
        Assert.Equal(all.Count, streamed.Count);
        for (int index = 0; index < all.Count; index++) Assert.Equal(all[index].Bytes, streamed[index].Bytes);
    }

    [Fact]
    public void HtmlImages_ExplicitScreenFullPageRemainsAContinuousScreenshot() {
        var document = HtmlConversionDocument.Parse(PagedImageHtml);
        var request = HtmlRenderRequest.Create(HtmlRenderIntentProfile.ScreenFullPage,
            HtmlRenderEncoder.Png, new HtmlRenderOptions {
                ViewportWidth = 240D, Margins = HtmlRenderMargins.All(0D), AllowSystemFontFallback = false
            });
        var actual = Assert.Single(document.RenderImages(request));
        var continuous = request.Options;
        continuous.Mode = HtmlRenderMode.Continuous;
        var expected = Assert.Single(document.ExportImages(OfficeImageExportFormat.Png, continuous));
        Assert.Equal(240, actual.Width);
        Assert.True(actual.Height < 120);
        Assert.Equal(expected.Bytes, actual.Bytes);
    }
}
