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
        byte[] bytes = document.ToImage(options, out OfficeConversionReport report);
        OfficeConversionResult<byte[]> asynchronous = await document.ToImageAsync(options);
        using var stream = new MemoryStream();
        OfficeConversionReport saved = await document.SaveAsImageAsync(stream, options);

        Assert.Equal(format.GetMimeType(), OfficeImageReader.Identify(bytes).MimeType);
        Assert.Equal(bytes, asynchronous.Value);
        Assert.Equal(bytes, stream.ToArray());
        Assert.Equal("HTML", report.Source);
        Assert.Equal(format.ToString(), report.Target);
        Assert.Equal(1, report.ItemsWritten);
        Assert.Equal(report.Target, asynchronous.Report.Target);
        Assert.Equal(report.Target, saved.Target);
        Assert.True(stream.CanWrite);
        Assert.Null(saved.OutputPath);
        Assert.Equal(1D, options.Scale);
    }

    [Fact]
    public void ImageConversionSelectsTheConfiguredPageBeforeEncoding() {
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            "<style>@page{size:100px 80px;margin:0}</style>" +
            "<div style='height:20px;background:red'>First</div>" +
            "<div style='break-before:page;height:20px;background:blue'>Second</div>");
        var options = new HtmlRenderOptions { Mode = HtmlRenderMode.Paged,
            Format = OfficeImageExportFormat.Svg, PageIndex = 1 };
        byte[] expected = document.ExportImage(OfficeImageExportFormat.Svg, options, pageIndex: 1).Bytes;
        Assert.Equal(expected, document.ToImage(options, out OfficeConversionReport report));
        Assert.Equal(1, report.ItemsWritten);
        Assert.Equal(1, options.PageIndex);
    }

    [Fact]
    public async Task CanceledImageConversionsDoNotWriteToCallerStreams() {
        HtmlConversionDocument document = HtmlConversionDocument.Parse("<p>Keep output empty</p>");
        using var stream = new MemoryStream();
        var cancellation = new CancellationToken(canceled: true);
        Assert.ThrowsAny<OperationCanceledException>(() => document.ToImage(cancellationToken: cancellation));
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            document.SaveAsImageAsync(stream, cancellationToken: cancellation));
        Assert.Equal(0, stream.Length);
        Assert.True(stream.CanWrite);
    }
}
