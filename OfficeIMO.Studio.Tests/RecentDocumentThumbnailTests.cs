using Avalonia.Controls;
using Avalonia.Headless;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Home;
using OfficeIMO.Studio.Features.Shell;

namespace OfficeIMO.Studio.Tests;

public sealed class RecentDocumentThumbnailTests {
    [Fact]
    public async Task FailedPreviewCanRetryWithoutMetadataChange() {
        using var app = TestAppBuilder.StartSession();
        await app.Dispatch(async () => {
            string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-preview-retry-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(root);
            try {
                string path = Path.Combine(root, "preview.pdf");
                byte[] valid = PdfDocument.Create(compose => compose.Page(page => page.Size(600D, 800D))).ToBytes();
                DateTime stamp = DateTime.UtcNow.AddMinutes(-5);
                File.WriteAllBytes(path, new byte[valid.Length]);
                File.SetLastWriteTimeUtc(path, stamp);
                string? fingerprint = RecentDocumentThumbnails.GetFingerprint(path);
                Assert.Null((await RecentDocumentThumbnails.GetAsync(path))?.Image);

                File.WriteAllBytes(path, valid);
                File.SetLastWriteTimeUtc(path, stamp);
                Assert.Equal(fingerprint, RecentDocumentThumbnails.GetFingerprint(path));
                Assert.NotNull((await RecentDocumentThumbnails.GetAsync(path))?.Image);
            } finally { Directory.Delete(root, recursive: true); }
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task ExistingRecentCardRefreshesAfterItsPdfChanges() {
        using var app = TestAppBuilder.StartSession();
        await app.Dispatch(async () => {
            string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-preview-refresh-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(root);
            try {
                string path = Path.Combine(root, "preview.pdf");
                PdfDocument.Create(compose => compose.Page(page => page.Size(600D, 800D))).Save(path);
                var recent = new RecentDocumentViewModel(path, DateTimeOffset.UtcNow);
                recent.EnsureThumbnail();
                using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(10));
                while (recent.Thumbnail is null) await Task.Delay(10, timeout.Token);
                var first = recent.Thumbnail;

                PdfDocument.Create(compose => {
                    compose.Page(page => page.Size(600D, 800D));
                    compose.Page(page => page.Size(600D, 800D));
                }).Save(path);
                File.SetLastWriteTimeUtc(path, DateTime.UtcNow.AddMinutes(1));
                recent.EnsureThumbnail();
                while (recent.Thumbnail is null || ReferenceEquals(first, recent.Thumbnail) ||
                       !string.Equals(recent.PageCountLabel, "2 pages", StringComparison.Ordinal))
                    await Task.Delay(10, timeout.Token);
                Assert.NotSame(first, recent.Thumbnail);
                using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null));
                model.RecentDocuments.Add(recent);
                var window = new Window { Width = 1000, Height = 950, Content = new HomeView { DataContext = model } };
                try {
                    window.Show();
                    window.UpdateLayout();
                    string? folder = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
                    if (!string.IsNullOrWhiteSpace(folder)) {
                        Directory.CreateDirectory(folder);
                        using var frame = window.CaptureRenderedFrame();
                        Assert.NotNull(frame);
                        frame.Save(Path.Combine(folder, "recent-card-refreshed.png"), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
                    }
                } finally { window.Close(); }
            } finally { Directory.Delete(root, recursive: true); }
            return true;
        }, CancellationToken.None);
    }

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
