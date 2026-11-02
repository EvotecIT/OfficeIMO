using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData(OfficeImageExportFormat.Png)]
    [InlineData(OfficeImageExportFormat.Jpeg)]
    [InlineData(OfficeImageExportFormat.Tiff)]
    [InlineData(OfficeImageExportFormat.Webp)]
    public async Task HtmlImageExport_EncodesTheRequestedRasterFormat(OfficeImageExportFormat format) {
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            "<div style='width:120px;height:48px;background:#336699;color:white'>Format marker</div>");
        var options = new HtmlRenderOptions {
            ViewportWidth = 180D,
            Margins = HtmlRenderMargins.All(8D)
        };

        OfficeImageExportResult result = document.ExportImage(format, options);
        OfficeImageExportResult asyncResult = await document.ExportImageAsync(format, options);

        Assert.Equal(format, result.Format);
        Assert.Equal(format, asyncResult.Format);
        Assert.Equal(format.GetMimeType(), OfficeImageReader.Identify(result.Bytes).MimeType);
        Assert.Equal(format.GetMimeType(), OfficeImageReader.Identify(asyncResult.Bytes).MimeType);
        Assert.Equal(result.Width, asyncResult.Width);
        Assert.Equal(result.Height, asyncResult.Height);
    }

    [Fact]
    public void HtmlRenderOptions_ClonePreservesRasterEncodingSettings() {
        var options = new HtmlRenderOptions();
        options.RasterEncoding.Jpeg.Quality = 63;
        options.RasterEncoding.Jpeg.Progressive = true;
        options.RasterEncoding.Tiff.Compression = OfficeTiffCompression.None;

        HtmlRenderOptions clone = options.Clone();

        Assert.NotSame(options.RasterEncoding, clone.RasterEncoding);
        Assert.NotSame(options.RasterEncoding.Jpeg, clone.RasterEncoding.Jpeg);
        Assert.NotSame(options.RasterEncoding.Tiff, clone.RasterEncoding.Tiff);
        Assert.Equal(63, clone.RasterEncoding.Jpeg.Quality);
        Assert.True(clone.RasterEncoding.Jpeg.Progressive);
        Assert.Equal(OfficeTiffCompression.None, clone.RasterEncoding.Tiff.Compression);
    }

    [Fact]
    public void HtmlImageExport_SingleAndBatchOperationsUseSharedFormats() {
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            "<style>@page{size:2in 1in;margin:8px}</style><p>Fluent HTML image export</p>");

        var imageOptions1 = new HtmlRenderOptions();
        imageOptions1.Mode = HtmlRenderMode.Paged;
        imageOptions1.PageIndex = 0;
        imageOptions1.Format = OfficeImageExportFormat.Jpeg;
        imageOptions1.RasterEncoding.Jpeg.Quality = 71;
        OfficeImageExportResult jpeg = document.ExportImage(OfficeImageExportFormat.Jpeg, imageOptions1, pageIndex: imageOptions1.PageIndex);
        var imageOptions2 = new HtmlRenderOptions();
        imageOptions2.Mode = HtmlRenderMode.Paged;
        imageOptions2.Format = OfficeImageExportFormat.Tiff;
        IReadOnlyList<OfficeImageExportResult> tiffPages = document.ExportImages(OfficeImageExportFormat.Tiff, imageOptions2);

        Assert.Equal("image/jpeg", OfficeImageReader.Identify(jpeg.Bytes).MimeType);
        Assert.NotEmpty(tiffPages);
        Assert.All(tiffPages, result => Assert.Equal("image/tiff", OfficeImageReader.Identify(result.Bytes).MimeType));
    }

    [Fact]
    public async Task HtmlImageExport_RasterConvenienceMethodsWriteCallerOwnedStreams() {
        HtmlConversionDocument document = HtmlConversionDocument.Parse("<p>Convenience methods</p>");
        using var jpeg = new MemoryStream();
        using var tiff = new MemoryStream();
        using var webp = new MemoryStream();

        OfficeImageExportResult jpegResult = document.SaveAsJpeg(jpeg);
        OfficeImageExportResult tiffResult = await document.SaveAsTiffAsync(tiff);
        OfficeImageExportResult webpResult = document.SaveAsWebp(webp);

        Assert.True(jpeg.CanWrite);
        Assert.True(tiff.CanWrite);
        Assert.True(webp.CanWrite);
        Assert.Equal("image/jpeg", OfficeImageReader.Identify(jpeg.ToArray()).MimeType);
        Assert.Equal("image/tiff", OfficeImageReader.Identify(tiff.ToArray()).MimeType);
        Assert.Equal("image/webp", OfficeImageReader.Identify(webp.ToArray()).MimeType);
        Assert.Equal(OfficeImageExportFormat.Jpeg, jpegResult.Format);
        Assert.Equal(OfficeImageExportFormat.Tiff, tiffResult.Format);
        Assert.Equal(OfficeImageExportFormat.Webp, webpResult.Format);
    }

    [Fact]
    public async Task HtmlSaveAsImageAsyncUsesTheResourceAwareRenderer() {
        bool resolverCalled = false;
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            "<link rel='stylesheet' href='https://assets.example.test/site.css'><p class='marker'>Async builder</p>");
        var options = new HtmlRenderOptions {
            ResourceResolver = (request, cancellationToken) => {
                cancellationToken.ThrowIfCancellationRequested();
                resolverCalled = true;
                return Task.FromResult<HtmlResolvedResource?>(
                    new HtmlResolvedResource(
                        System.Text.Encoding.UTF8.GetBytes(".marker{color:#336699}"),
                        "text/css"));
            }
        };
        using var output = new MemoryStream();

        var imageOptions3 = options.Clone();
        imageOptions3.Format = OfficeImageExportFormat.Png;
        OfficeConversionReport result = await document
            .SaveAsImageAsync(output, imageOptions3);

        Assert.True(resolverCalled);
        Assert.Equal("Png", result.Target);
        Assert.Equal(1, result.ItemsWritten);
        Assert.True(output.CanWrite);
        Assert.Equal("image/png", OfficeImageReader.Identify(output.ToArray()).MimeType);
    }

    [Fact]
    public void HtmlImageExport_RejectsOverflowingScaleBeforeEncoding() {
        HtmlConversionDocument document = HtmlConversionDocument.Parse("<p>Overflow guard</p>");
        var options = new HtmlRenderOptions { Scale = double.MaxValue };

        Assert.Throws<InvalidOperationException>(() =>
            document.ExportImage(OfficeImageExportFormat.Svg, options));
    }

    [Fact]
    public void HtmlImageExport_StrictOmissionPolicyRejectsUndecodableGeneratedImage() {
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            "<style>.marker::before{content:url('data:image/png;base64,AA==')}</style><p class='marker'>Body</p>");
        var options = new HtmlRenderOptions {
            Policy = new OfficeImageExportPolicy {
                RequireNoOmissions = true
            }
        };

        OfficeImageExportPolicyException exception = Assert.Throws<OfficeImageExportPolicyException>(
            () => document.ExportImage(OfficeImageExportFormat.Png, options));

        Assert.Contains(
            exception.Diagnostics,
            diagnostic =>
                diagnostic.Code == HtmlPdfDiagnosticCodes.ImagePayloadOmitted &&
                diagnostic.LossKind == OfficeConversionLossKind.Omission);
    }

    [Fact]
    public void HtmlImageExport_BatchSavePreservesCallerRenderOptions() {
        HtmlConversionDocument document = HtmlConversionDocument.Parse("<p>Configured batch</p>");
        var options = new HtmlRenderOptions {
            ViewportWidth = 222D,
            ViewportHeight = 111D,
            Scale = 0.5D,
            Margins = HtmlRenderMargins.All(4D)
        };
        string folder = Path.Combine(Path.GetTempPath(), "OfficeIMO-" + Guid.NewGuid().ToString("N"));
        try {
            options.Mode = HtmlRenderMode.Continuous;
            OfficeConversionReport report = document.SaveAsImages(folder, options);
            Assert.Equal(1, report.ItemsWritten);
            Assert.Equal(222D, options.ViewportWidth);
            Assert.Equal(0.5D, options.Scale);
            OfficeImageInfo result = OfficeImageReader.Identify(File.ReadAllBytes(Path.Combine(folder, "Page 1.png")));
            Assert.Equal(111, result.Width);
            Assert.Equal(56, result.Height);
            Assert.True(File.Exists(Path.Combine(folder, "Page 1.png")));
        } finally {
            if (Directory.Exists(folder)) Directory.Delete(folder, recursive: true);
        }
    }
}
