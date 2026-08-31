using OfficeIMO.Studio.Features.Reader;
using OfficeIMO.Pdf;

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

    [Fact]
    public async Task SearchReportsPerPageProgressAndCancelsBeforeTheNextPage() {
        string path = System.IO.Path.Combine(
            System.IO.Path.GetTempPath(),
            $"officeimo-studio-search-{Guid.NewGuid():N}.pdf");
        PdfDocument.Create(compose => {
            compose.Page(page => page.Size(500, 700));
            compose.Page(page => page.Size(500, 700));
        }).Save(path);
        using var cancellation = new CancellationTokenSource();
        var progress = new InlineProgress<double>(fraction => {
            if (fraction >= 0.5D) cancellation.Cancel();
        });

        try {
            PdfDocumentSession session = await PdfDocumentSession.OpenAsync(path, CancellationToken.None);

            await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
                session.SearchAsync("needle", cancellation.Token, progress));

            Assert.Equal(0.5D, Assert.Single(progress.Values), precision: 3);
        } finally {
            File.Delete(path);
        }
    }

    private static string GetFixturePath() =>
        System.IO.Path.Combine(AppContext.BaseDirectory, "Fixtures", "openpreserve-pdfa1b-text.pdf");

    private sealed class InlineProgress<T>(Action<T> report) : IProgress<T> {
        internal List<T> Values { get; } = new();

        public void Report(T value) {
            Values.Add(value);
            report(value);
        }
    }
}
