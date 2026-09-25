using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Home;

namespace OfficeIMO.Studio.Tests;

public sealed class RecentDocumentThumbnailTests {
    [Fact]
    public async Task CancelledPreviewCanBeRequestedAgain() {
        using var app = TestAppBuilder.StartSession();
        await app.Dispatch(async () => {
            string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-preview-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(root);
            try {
                string path = Path.Combine(root, "preview.pdf");
                PdfDocument.Create(compose => compose.Page(page => page.Size(600D, 800D)
                    .Canvas(canvas => canvas.Text("Preview", 30D, 100D, 100D, 30D, fontSize: 12D)))).Save(path);
                using var cancelled = new CancellationTokenSource();
                cancelled.Cancel();

                await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
                    RecentDocumentThumbnails.GetAsync(path, cancelled.Token));
                RecentDocumentPreview preview = Assert.IsType<RecentDocumentPreview>(
                    await RecentDocumentThumbnails.GetAsync(path));
                Assert.Equal(1, preview.PageCount);
                Assert.NotNull(preview.Image);
            } finally { Directory.Delete(root, recursive: true); }
            return true;
        }, CancellationToken.None);
    }
}
