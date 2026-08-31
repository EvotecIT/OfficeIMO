using OfficeIMO.Studio.Features.Reader;

namespace OfficeIMO.Studio.Tests;

public sealed class PdfPageViewModelTests {
    private static readonly byte[] TinyPng = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");

    [Fact]
    public async Task NewerZoomGenerationWinsOverSlowRender() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            int calls = 0;
            var firstEntered = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
            var releaseFirst = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);

            using var coordinator = new PageRenderCoordinator(async (page, scale, cancellationToken) => {
                int call = Interlocked.Increment(ref calls);
                if (call == 1) {
                    firstEntered.SetResult();
                    await releaseFirst.Task;
                    cancellationToken.ThrowIfCancellationRequested();
                }

                return new PdfRenderedPage(page, scale, TinyPng, 1, 1, TimeSpan.Zero, Array.Empty<string>());
            });
            using var viewModel = new PdfPageViewModel(1, 612, 792, 0, 1D, coordinator);

            viewModel.AttachToViewport();
            await firstEntered.Task;
            viewModel.SetZoom(2D);
            releaseFirst.SetResult();

            await WaitUntilAsync(() => calls >= 2 && viewModel.PageImage is not null && !viewModel.IsRendering);

            Assert.Equal(2, calls);
            Assert.Equal(2D, viewModel.RenderedScale);
            Assert.NotNull(viewModel.PageImage);
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task DetachingReleasesDecodedBitmapButKeepsEncodedCache() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            using var coordinator = new PageRenderCoordinator((page, scale, _) =>
                Task.FromResult(new PdfRenderedPage(page, scale, TinyPng, 1, 1, TimeSpan.Zero, Array.Empty<string>())));
            using var viewModel = new PdfPageViewModel(1, 612, 792, 0, 1D, coordinator);

            viewModel.AttachToViewport();
            await WaitUntilAsync(() => viewModel.PageImage is not null && !viewModel.IsRendering);
            viewModel.DetachFromViewport();

            Assert.Null(viewModel.PageImage);
            Assert.Equal(1, coordinator.CachedEntryCount);
            return true;
        }, CancellationToken.None);
    }

    private static async Task WaitUntilAsync(Func<bool> condition) {
        using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(5));
        while (!condition()) {
            await Task.Delay(10, timeout.Token);
        }
    }
}
