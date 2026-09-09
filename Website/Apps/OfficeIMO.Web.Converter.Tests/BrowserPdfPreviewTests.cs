using OfficeIMO.Pdf;
using OfficeIMO.Web.Converter.Services;
using Xunit;

namespace OfficeIMO.Web.Converter.Tests;

public sealed class BrowserPdfPreviewTests {
    [Fact]
    public void PreviewRendersActualPageAsBoundedPng() {
        byte[] source = PdfDocument.Create(document => document.Content(content => content.H1("Preview without a PDF plugin"))).ToBytes();
        var preview = new BrowserPdfPreview(source);

        PdfPageRenderResult page = preview.Render(1);

        Assert.Equal(1, preview.PageCount);
        Assert.True(page.Succeeded);
        Assert.Equal(new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 }, page.Bytes![..8]);
        Assert.InRange(page.Width, 1, 1280);
        Assert.InRange(page.Height, 1, 1280);
        Assert.Throws<ArgumentOutOfRangeException>(() => preview.Render(2));
    }

    [Fact]
    public void CancelledRenderDoesNotProduceAnImage() {
        byte[] source = PdfDocument.Create(document => document.Content(content => content.H1("Cancelled preview"))).ToBytes();
        var preview = new BrowserPdfPreview(source);
        Assert.ThrowsAny<OperationCanceledException>(() => preview.Render(1, new CancellationToken(true)));
    }

    [Fact]
    public void InvalidPdfIsRejectedWithoutMutatingSource() {
        byte[] source = [1, 2, 3];
        Assert.ThrowsAny<Exception>(() => new BrowserPdfPreview(source));
        Assert.Equal(new byte[] { 1, 2, 3 }, source);
    }
}
