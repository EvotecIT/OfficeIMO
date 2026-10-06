using OfficeIMO.Studio.Features.Organizer;
using OfficeIMO.Studio.Features.Reader;

namespace OfficeIMO.Studio.Tests;

public sealed class PdfInitialRasterScalingTests {
    private static readonly byte[] TinyPng = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");

    [Theory]
    [InlineData(false, false, 1D, 2D, 2D)]
    [InlineData(false, true, 1D, 2D, 2D)]
    [InlineData(true, false, 1D, 2D, 0.5D)]
    [InlineData(true, false, 4D, 5D, 1D)]
    public async Task InitialRasterUsesPresentationChangedDuringSceneLoading(bool organizer, bool changeZoom, double initialScale, double updatedScale, double expected) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var entered = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
            var sceneReady = new TaskCompletionSource<PdfPageScene>(TaskCreationOptions.RunContinuationsAsynchronously);
            using var scenes = new PageSceneCoordinator((_, _) => {
                entered.TrySetResult();
                return sceneReady.Task;
            });
            var scales = new List<double>();
            using var renders = new PageRenderCoordinator((page, scale, _) => {
                scales.Add(scale);
                return Task.FromResult(new PdfRenderedPage(page, scale, TinyPng, 1, 1, TimeSpan.Zero, []));
            });
            using var reader = organizer ? null : new PdfPageViewModel(1, 612D, 792D, 0, 1D, scenes, renders);
            using var thumbnail = organizer ? new PdfOrganizerPageViewModel(1, 612D, 792D, 0, scenes, renders) : null;
            reader?.SetRenderScaling(initialScale);
            thumbnail?.SetRenderScaling(initialScale);
            reader?.AttachToViewport();
            thumbnail?.Attach();
            await entered.Task.WaitAsync(TimeSpan.FromSeconds(10));
            if (changeZoom) reader!.SetZoom(updatedScale);
            else if (organizer) thumbnail!.SetRenderScaling(updatedScale);
            else reader!.SetRenderScaling(updatedScale);
            sceneReady.SetResult(TestPdfPageScenes.Create(requiresRasterFallback: true));
            using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(10));
            while (organizer ? thumbnail!.FallbackImage is null || thumbnail.IsLoading : reader!.PageImage is null || reader.IsRendering)
                await Task.Delay(10, timeout.Token);
            Assert.Equal(expected, Assert.Single(scales));
            return true;
        }, default);
    }
}
