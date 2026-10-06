using OfficeIMO.Studio.Features.Organizer;
using OfficeIMO.Studio.Features.Reader;

namespace OfficeIMO.Studio.Tests;

public sealed class PdfInitialRasterScalingTests {
    private static readonly byte[] TinyPng = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");

    [Theory]
    [InlineData(false, false, 2D)]
    [InlineData(false, true, 2D)]
    [InlineData(true, false, 0.5D)]
    public async Task InitialRasterUsesPresentationChangedDuringSceneLoading(bool organizer, bool changeZoom, double expected) {
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
            reader?.AttachToViewport();
            thumbnail?.Attach();
            await entered.Task.WaitAsync(TimeSpan.FromSeconds(10));
            if (changeZoom) reader!.SetZoom(2D);
            else if (organizer) thumbnail!.SetRenderScaling(2D);
            else reader!.SetRenderScaling(2D);
            sceneReady.SetResult(TestPdfPageScenes.Create(requiresRasterFallback: true));
            using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(10));
            while (organizer ? thumbnail!.FallbackImage is null || thumbnail.IsLoading : reader!.PageImage is null || reader.IsRendering)
                await Task.Delay(10, timeout.Token);
            Assert.Equal(expected, Assert.Single(scales));
            return true;
        }, default);
    }
}
