using System.Text;
using System.Threading;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfPageImageRendererBatchTests {
    [Fact]
    public void RenderPages_RendersCallerOrderedSvgRangeWithPerPageReports() {
        byte[] pdf = BuildTwoPagePdf();

        IReadOnlyList<PdfPageRenderResult> results = PdfDocument.Open(pdf).Read.RenderPages(
            PdfPageSelection.From(2, 1),
            new PdfPageRenderOptions { Format = PdfPageRenderFormat.Svg, Scale = 2D });

        Assert.Equal(new[] { 2, 1 }, results.Select(static result => result.PageNumber));
        Assert.All(results, result => {
            Assert.True(result.Succeeded, string.Join(" ", result.Diagnostics));
            Assert.StartsWith("<svg", Encoding.UTF8.GetString(result.Bytes!), StringComparison.Ordinal);
            Assert.True(result.Width > 1000);
            Assert.True(result.Height > 1500);
            Assert.True(result.Elapsed >= TimeSpan.Zero);
        });
    }

    [Fact]
    public void RenderPages_SupportsThumbnailsCancellationAndExplicitLimits() {
        byte[] pdf = BuildTwoPagePdf();
        IReadOnlyList<PdfPageRenderResult> thumbnails = PdfPageImageRenderer.RenderPages(
            pdf,
            "1-2",
            new PdfPageRenderOptions { ThumbnailMaxDimension = 96 });

        Assert.All(thumbnails, result => {
            Assert.True(result.Succeeded, string.Join(" ", result.Diagnostics));
            Assert.True(Math.Max(result.Width, result.Height) <= 96);
            Assert.Equal(new byte[] { 137, 80, 78, 71 }, result.Bytes!.Take(4));
        });

        PdfReadLimitException pageLimit = Assert.Throws<PdfReadLimitException>(() =>
            PdfPageImageRenderer.RenderPages(pdf, options: new PdfPageRenderOptions { MaxPages = 1 }));
        Assert.Equal(PdfReadLimitKind.RenderPages, pageLimit.Kind);

        PdfPageRenderResult pixelFailure = Assert.Single(PdfPageImageRenderer.RenderPages(
            pdf,
            PdfPageSelection.From(1),
            new PdfPageRenderOptions { MaxPixelsPerPage = 10 }));
        Assert.False(pixelFailure.Succeeded);
        Assert.Contains("render pixel count", Assert.Single(pixelFailure.Diagnostics), StringComparison.OrdinalIgnoreCase);

        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() =>
            PdfPageImageRenderer.RenderPages(pdf, cancellationToken: cancellation.Token));
    }

    private static byte[] BuildTwoPagePdf() => PdfDocument.Create()
        .Paragraph(paragraph => paragraph.Text("Batch page one"))
        .PageBreak()
        .Paragraph(paragraph => paragraph.Text("Batch page two"))
        .ToBytes();
}
