using OfficeIMO.Studio.Features.Reader;

namespace OfficeIMO.Studio.Tests;

public sealed class PageRenderCoordinatorTests {
    [Fact]
    public async Task ReusesCachedPageForEquivalentScaleBucket() {
        int renderCount = 0;
        using var coordinator = new PageRenderCoordinator(
            (page, scale, _) => {
                renderCount++;
                return Task.FromResult(CreatePage(page, scale, 32));
            });

        PdfRenderedPage first = await coordinator.GetPageAsync(1, 1D, CancellationToken.None);
        PdfRenderedPage second = await coordinator.GetPageAsync(1, 1.02D, CancellationToken.None);

        Assert.Same(first, second);
        Assert.Equal(1, renderCount);
        Assert.Equal(1, coordinator.CachedEntryCount);
    }

    [Fact]
    public async Task EvictsLeastRecentlyUsedEntriesWithinConfiguredBounds() {
        using var coordinator = new PageRenderCoordinator(
            (page, scale, _) => Task.FromResult(CreatePage(page, scale, 40)),
            maximumEntries: 2,
            maximumBytes: 90);

        await coordinator.GetPageAsync(1, 1D, CancellationToken.None);
        await coordinator.GetPageAsync(2, 1D, CancellationToken.None);
        await coordinator.GetPageAsync(1, 1D, CancellationToken.None);
        await coordinator.GetPageAsync(3, 1D, CancellationToken.None);

        Assert.Equal(2, coordinator.CachedEntryCount);
        Assert.Equal(80, coordinator.CachedByteCount);
    }

    [Fact]
    public async Task DisposalCancelsActiveRender() {
        CancellationToken cancellationToken = CancellationToken.None;
        var entered = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var cancelled = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var coordinator = new PageRenderCoordinator(async (_, _, cancellationToken) => {
            entered.SetResult();
            try {
                await Task.Delay(Timeout.InfiniteTimeSpan, cancellationToken);
                throw new InvalidOperationException("The render should have been cancelled.");
            } catch (OperationCanceledException) {
                cancelled.SetResult();
                throw;
            }
        });

        Task<PdfRenderedPage> render = coordinator.GetPageAsync(1, 1D, CancellationToken.None);
        await entered.Task;
        coordinator.Dispose();

        await cancelled.Task.WaitAsync(TimeSpan.FromSeconds(2), cancellationToken);
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => render);
    }

    private static PdfRenderedPage CreatePage(int page, double scale, int bytes) =>
        new(page, scale, new byte[bytes], 10, 10, TimeSpan.Zero, Array.Empty<string>());
}
