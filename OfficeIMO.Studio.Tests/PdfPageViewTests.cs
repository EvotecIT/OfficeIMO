using Avalonia;
using Avalonia.Controls;
using OfficeIMO.Studio.Features.Reader;

namespace OfficeIMO.Studio.Tests;

public sealed class PdfPageViewTests {
    private static readonly byte[] TinyPng = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");

    [Fact]
    public async Task AttachingAfterDataContextStartsPageRendering() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            int renderCalls = 0;
            using var coordinator = new PageRenderCoordinator((page, scale, _) => {
                Interlocked.Increment(ref renderCalls);
                return Task.FromResult(
                    new PdfRenderedPage(page, scale, TinyPng, 1, 1, TimeSpan.Zero, Array.Empty<string>()));
            });
            using var sceneCoordinator = new PageSceneCoordinator((page, _) =>
                Task.FromResult(TestPdfPageScenes.Create(page, requiresRasterFallback: true)));
            using var viewModel = new PdfPageViewModel(1, 612, 792, 0, 1D, sceneCoordinator, coordinator);
            var view = new PdfPageView { DataContext = viewModel };
            var window = new Window { Content = view };

            try {
                window.Show();
                window.Measure(new Size(800, 600));
                window.Arrange(new Rect(0, 0, 800, 600));
                await WaitUntilAsync(() => viewModel.PageImage is not null && !viewModel.IsRendering);

                Assert.Equal(1, renderCalls);
                Assert.NotNull(viewModel.PageImage);
            } finally {
                window.Close();
            }

            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task PendingRedactionAreaBindsToThePageCanvas() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(() => {
            using var renderCoordinator = new PageRenderCoordinator((page, scale, _) =>
                Task.FromResult(new PdfRenderedPage(page, scale, TinyPng, 1, 1, TimeSpan.Zero, Array.Empty<string>())));
            using var sceneCoordinator = new PageSceneCoordinator((page, _) =>
                Task.FromResult(TestPdfPageScenes.Create(page)));
            using var viewModel = new PdfPageViewModel(1, 612, 792, 0, 1D, sceneCoordinator, renderCoordinator) {
                PendingRedactionArea = new Rect(42D, 64D, 180D, 36D)
            };
            var view = new PdfPageView { DataContext = viewModel };
            var window = new Window { Content = view };

            try {
                window.Show();
                window.Measure(new Size(800, 900));
                window.Arrange(new Rect(0, 0, 800, 900));

                PdfPageCanvas canvas = Assert.IsType<PdfPageCanvas>(view.FindControl<PdfPageCanvas>("PageCanvas"));
                Assert.Equal(viewModel.PendingRedactionArea, canvas.PendingRedactionArea);
            } finally {
                window.Close();
            }

            return Task.CompletedTask;
        }, CancellationToken.None);
    }

    private static async Task WaitUntilAsync(Func<bool> condition) {
        using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(5));
        while (!condition()) await Task.Delay(10, timeout.Token);
    }
}
