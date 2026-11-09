using System.Linq;
using OfficeIMO.Html.Pdf;
using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData(OfficeImageExportFormat.Png)]
    [InlineData(OfficeImageExportFormat.Jpeg)]
    [InlineData(OfficeImageExportFormat.Tiff)]
    [InlineData(OfficeImageExportFormat.Webp)]
    [InlineData(OfficeImageExportFormat.Svg)]
    public async Task ImageConversionUsesConfiguredEncodingAndReturnsSharedEvidence(OfficeImageExportFormat format) {
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            "<div style='width:40px;height:24px;background:#336699'>Preview</div>");
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Continuous, Format = format,
            ViewportWidth = 80, ViewportHeight = 40, Margins = HtmlRenderMargins.All(0)
        };
        byte[] bytes = document.ToImage(options, out OfficeConversionReport report).Bytes;
        OfficeImageExportResult asynchronous = await document.ToImageAsync(options);
        using var stream = new MemoryStream();
        OfficeConversionReport saved = await document.SaveAsImageAsync(stream, options);

        Assert.Equal(format.GetMimeType(), OfficeImageReader.Identify(bytes).MimeType);
        Assert.Equal(bytes, asynchronous.Bytes);
        Assert.Equal(bytes, stream.ToArray());
        Assert.Equal("HTML", report.Source);
        Assert.Equal(format.ToString(), report.Target);
        Assert.Equal(1, report.ItemsWritten);
        Assert.Equal(report.Target, asynchronous.CreateReport().Target);
        Assert.Equal(report.Target, saved.Target);
        Assert.True(stream.CanWrite);
        Assert.Null(saved.OutputPath);
        Assert.Equal(1D, options.Scale);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task SharedImageSelectorPreservesPageOrderRepetitionsAndSavedCount(bool asynchronous) {
        var document = HtmlConversionDocument.Parse("<style>@page{size:100px 80px;margin:0}</style>"
            + "<div style='height:20px;background:red'>First</div>"
            + "<div style='break-before:page;height:20px;background:blue'>Second</div>");
        var options = new HtmlRenderOptions { Mode = HtmlRenderMode.Paged, Format = OfficeImageExportFormat.Svg,
            PageSelector = OfficePageSelector.Parse("2,1,2") };
        var images = asynchronous ? await document.ToImagesAsync(options) : document.ToImages(options);
        Assert.Equal(new[] { "Page 2", "Page 1", "Page 2" }, images.Select(image => image.Name));
        Assert.Contains("#0000ff", System.Text.Encoding.UTF8.GetString(images[0].Bytes), StringComparison.OrdinalIgnoreCase);
        Assert.Contains("#ff0000", System.Text.Encoding.UTF8.GetString(images[1].Bytes), StringComparison.OrdinalIgnoreCase);
        string folder = Path.Combine(Path.GetTempPath(), "officeimo-html-selector-" + Guid.NewGuid().ToString("N"));
        try {
            OfficeConversionReport report = asynchronous ? await document.SaveAsImagesAsync(folder, options) : document.SaveAsImages(folder, options);
            Assert.Equal(3, report.ItemsWritten); Assert.Equal(3, Directory.GetFiles(folder, "*.svg").Length);
        } finally { if (Directory.Exists(folder)) Directory.Delete(folder, true); }
    }

    [Fact]
    public void ImageConversionSelectsTheConfiguredPageBeforeEncoding() {
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            "<style>@page{size:100px 80px;margin:0}</style>" +
            "<div style='height:20px;background:red'>First</div>" +
            "<div style='break-before:page;height:20px;background:blue'>Second</div>");
        var options = new HtmlRenderOptions { Mode = HtmlRenderMode.Paged,
            Format = OfficeImageExportFormat.Svg, PageIndex = 1 };
        byte[] expected = document.ToImage(OfficeImageExportFormat.Svg, options, pageIndex: 1).Bytes;
        Assert.Equal(expected, document.ToImage(options, out OfficeConversionReport report).Bytes);
        Assert.Equal(1, report.ItemsWritten);
        Assert.Equal(1, options.PageIndex);
    }

    [Fact]
    public async Task CanceledImageConversionsDoNotWriteToCallerStreams() {
        HtmlConversionDocument document = HtmlConversionDocument.Parse("<p>Keep output empty</p>");
        using var stream = new MemoryStream();
        var cancellation = new CancellationToken(canceled: true);
        Assert.ThrowsAny<OperationCanceledException>(() => document.ToImage(cancellationToken: cancellation).Bytes);
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            document.SaveAsImageAsync(stream, cancellationToken: cancellation));
        Assert.Equal(0, stream.Length);
        Assert.True(stream.CanWrite);
    }
}
