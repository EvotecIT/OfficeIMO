using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfScanPreviewTests {
    [Fact]
    public async Task PreviewOwnsItsPixelsAndCreatesOnlyTheReviewedRegion() {
        var document = Source();
        byte[] source = document.ToBytes();
        var options = new PdfOcrMergeOptions {
            Dpi = 72, Regions = new[] { new PdfOcrPageRegion(2, .25, .25, .5, .5) },
            ReadOptions = new() { PageSelection = PdfPageSelection.From(1, 2) }
        };
        Assert.Equal(new[] { 2 }, options.GetSelectedPages(2));
        var preview = await document.PreviewScanAsync(2, options);
        Assert.Equal(100, preview.Width); Assert.Equal(50, preview.Height);
        byte[] originalPixels = preview.GetSourcePng(), preparedPixels = preview.GetPreparedPng();
        byte[] disposableSource = preview.GetSourcePng(), disposablePrepared = preview.GetPreparedPng();
        Array.Clear(disposableSource, 0, disposableSource.Length);
        Array.Clear(disposablePrepared, 0, disposablePrepared.Length);
        Assert.Equal(originalPixels, preview.GetSourcePng());
        Assert.Equal(preparedPixels, preview.GetPreparedPng());
        var output = preview.CreateImagePdf();
        Assert.Equal(1, output.Inspect().PageCount);
        Assert.Equal(100, output.Render.Drawing(1).Width);
        Assert.Equal(50, output.Render.Drawing(1).Height);
        Assert.True(string.IsNullOrWhiteSpace(output.Reader.Text()));
        Assert.Equal(source, document.ToBytes());
    }

    [Fact]
    public async Task PreviewRejectsExcludedPagesInvalidRegionsBudgetsAndCancellation() {
        var document = Source();
        await Assert.ThrowsAsync<ArgumentException>(() => document.PreviewScanAsync(1,
            new() { ReadOptions = new() { PageSelection = PdfPageSelection.From(2) } }));
        await Assert.ThrowsAsync<ArgumentException>(() => document.PreviewScanAsync(1,
            new() { Regions = new[] { new PdfOcrPageRegion(3, 0, 0, 1, 1) } }));
        await Assert.ThrowsAsync<PdfReadLimitException>(() => document.PreviewScanAsync(1,
            new() { Dpi = 72, MaxPixelsPerPage = 1 }));
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => document.PreviewScanAsync(1, cancellationToken: cancellation.Token));
    }

    private static PdfDocument Source() => PdfDocument.Create(compose => {
        compose.Page(page => page.Size(200, 100).Margin(0).Content(c => c.Item(i => i.Paragraph(t => t.Text("First source page")))));
        compose.Page(page => page.Size(200, 100).Margin(0).Content(c => c.Item(i => i.Paragraph(t => t.Text("Second source page")))));
    });
}
