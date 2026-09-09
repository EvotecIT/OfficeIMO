using OfficeIMO.Pdf;
using OfficeIMO.Web.Converter.Services;
using Xunit;

namespace OfficeIMO.Web.Converter.Tests;

public sealed class BrowserPdfPreviewTests {
    [Fact]
    public void ViewingPreviewSupportsRestrictedPdfWithoutEnablingExtraction() {
        // Independent pypdf 6.16.2 AES-256 fixture: empty user password and no extraction permissions.
        byte[] source = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "samples", "view-only.pdf"));
        var preview = new BrowserPdfPreview(source);
        Assert.True(preview.Render(1).Succeeded);
        Assert.Throws<PdfPermissionDeniedException>(() => PdfDocument.Load(source).Read());
    }

    [Fact]
    public void PageCountDoesNotDecodePageContentAndHonorsCancellation() {
        byte[] source = PdfDocument.Create(document => {
            document.Page(page => page.Content(content => content.Text("First page content")));
            document.Page(page => page.Content(content => content.Text("Second page content")));
        }).ToBytes();
        var document = PdfDocument.Load(source, new PdfLoadOptions { Limits = new PdfReadLimits { MaxPageContentBytes = 1 } });
        PdfDocumentViewInfo geometry = document.InspectGeometryForViewing();
        Assert.Equal(2, geometry.PageCount);
        Assert.True(geometry.CanExtractContent);
        Assert.Null(geometry.LogicalContent);
        Assert.Throws<PdfReadLimitException>(() => document.Render.DisplayPage(1));
        Assert.ThrowsAny<OperationCanceledException>(() => new BrowserPdfPreview(source, new CancellationToken(true)));
        var preview = new BrowserPdfPreview(source);
        Assert.True(preview.Render(2).Succeeded);
        Assert.True(preview.Render(1).Succeeded);
    }

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
