using System.Globalization;
using OfficeIMO.Studio.Features.Reader;
using OfficeIMO.Studio.Infrastructure.Localization;

namespace OfficeIMO.Studio.Tests;

public sealed class PdfPageViewModelTests {
    private static readonly byte[] TinyPng = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");

    [Fact]
    public void PageLabelUsesTheConfiguredCulture() {
        using var coordinator = new PageRenderCoordinator((page, scale, _) =>
            Task.FromResult(new PdfRenderedPage(page, scale, TinyPng, 1, 1, TimeSpan.Zero, Array.Empty<string>())));
        using var sceneCoordinator = new PageSceneCoordinator((page, _) =>
            Task.FromResult(TestPdfPageScenes.Create(page)));
        var localizer = new StudioLocalizer(CultureInfo.GetCultureInfo(StudioCultureCatalog.PseudoCulture));
        using var viewModel = new PdfPageViewModel(2, 612, 792, 0, 1D, sceneCoordinator, coordinator, localizer);

        Assert.StartsWith("⟦", viewModel.PageLabel, StringComparison.Ordinal);
        Assert.Contains("2", viewModel.PageLabel, StringComparison.Ordinal);
    }

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
            using var sceneCoordinator = new PageSceneCoordinator((page, _) =>
                Task.FromResult(TestPdfPageScenes.Create(page, requiresRasterFallback: true)));
            using var viewModel = new PdfPageViewModel(1, 612, 792, 0, 1D, sceneCoordinator, coordinator);

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
            using var sceneCoordinator = new PageSceneCoordinator((page, _) =>
                Task.FromResult(TestPdfPageScenes.Create(page, requiresRasterFallback: true)));
            using var viewModel = new PdfPageViewModel(1, 612, 792, 0, 1D, sceneCoordinator, coordinator);

            viewModel.AttachToViewport();
            await WaitUntilAsync(() => viewModel.PageImage is not null && !viewModel.IsRendering);
            viewModel.DetachFromViewport();

            Assert.Null(viewModel.PageImage);
            Assert.Null(viewModel.Scene);
            Assert.Equal(1, sceneCoordinator.CachedEntryCount);
            Assert.Equal(1, coordinator.CachedEntryCount);
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task SupportedRetainedSceneDoesNotInvokeRasterRenderer() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            int renderCalls = 0;
            using var renderCoordinator = new PageRenderCoordinator((page, scale, _) => {
                Interlocked.Increment(ref renderCalls);
                return Task.FromResult(new PdfRenderedPage(page, scale, TinyPng, 1, 1, TimeSpan.Zero, Array.Empty<string>()));
            });
            using var sceneCoordinator = new PageSceneCoordinator((page, _) =>
                Task.FromResult(TestPdfPageScenes.Create(page)));
            using var viewModel = new PdfPageViewModel(1, 612, 792, 0, 1D, sceneCoordinator, renderCoordinator);

            viewModel.AttachToViewport();
            await WaitUntilAsync(() => viewModel.Scene is not null && !viewModel.IsRendering);

            Assert.Equal(0, renderCalls);
            Assert.Null(viewModel.PageImage);
            Assert.False(viewModel.Scene!.RequiresRasterFallback);
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task RenderingDiagnosticsExposeTheActualReasonAndFullDetails() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            using var renderCoordinator = new PageRenderCoordinator((page, scale, _) =>
                Task.FromResult(new PdfRenderedPage(page, scale, TinyPng, 1, 1, TimeSpan.Zero, Array.Empty<string>())));
            using var sceneCoordinator = new PageSceneCoordinator((page, _) =>
                Task.FromResult(TestPdfPageScenes.Create(
                    page,
                    diagnostics: ["Embedded font fallback is required.", "Advanced text metrics are retained."])));
            using var viewModel = new PdfPageViewModel(1, 612, 792, 0, 1D, sceneCoordinator, renderCoordinator);

            viewModel.AttachToViewport();
            await WaitUntilAsync(() => viewModel.Scene is not null && !viewModel.IsRendering);

            Assert.Equal("Embedded font fallback is required.", viewModel.DiagnosticsSummary);
            Assert.Contains("Advanced text metrics are retained.", viewModel.DiagnosticsDetails, StringComparison.Ordinal);
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
