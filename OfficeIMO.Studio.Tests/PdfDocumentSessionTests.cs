using OfficeIMO.Studio.Features.Reader;

namespace OfficeIMO.Studio.Tests;

public sealed class PdfDocumentSessionTests {
    [Fact]
    public async Task OpensExistingPdfAndRendersRequestedPage() {
        string path = GetFixturePath();
        CancellationToken cancellationToken = CancellationToken.None;

        PdfDocumentSession session = await PdfDocumentSession.OpenAsync(path, cancellationToken);
        PdfRenderedPage page = await session.RenderPageAsync(1, 1D, cancellationToken);

        Assert.Equal(System.IO.Path.GetFullPath(path), session.Path);
        Assert.NotEmpty(session.Pages);
        Assert.Equal(1, page.PageNumber);
        Assert.True(page.Bytes.Length > 8);
        Assert.Equal(new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 }, page.Bytes[..8]);
        Assert.True(page.PixelWidth > 0);
        Assert.True(page.PixelHeight > 0);
    }

    [Fact]
    public async Task RejectsNonPdfInputBeforeOpening() {
        string path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"officeimo-studio-{Guid.NewGuid():N}.txt");
        CancellationToken cancellationToken = CancellationToken.None;
        await File.WriteAllTextAsync(path, "not a PDF", cancellationToken);
        try {
            NotSupportedException error = await Assert.ThrowsAsync<NotSupportedException>(
                () => PdfDocumentSession.OpenAsync(path, cancellationToken));
            Assert.Contains("currently opens PDF", error.Message, StringComparison.OrdinalIgnoreCase);
        } finally {
            File.Delete(path);
        }
    }

    private static string GetFixturePath() =>
        System.IO.Path.Combine(AppContext.BaseDirectory, "Fixtures", "openpreserve-pdfa1b-text.pdf");
}
