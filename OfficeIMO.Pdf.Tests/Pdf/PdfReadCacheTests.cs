using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfReadCacheTests {
    [Fact]
    public Task CancelledCompositionWaitPreservesTheSharedFalseResult() =>
        VerifyCancelledWait(false);

    [Fact]
    public Task CancelledProfileWaitPreservesTheSharedProfile() {
        Assert.True(OfficeIccColorProfile.TryCreate(IccMabTestProfiles.CreateRgbXyz16WithDistinctOutputIntents(), out OfficeIccColorProfile? profile));
        return VerifyCancelledWait(profile!);
    }

    [Fact]
    public void CancelledInitializationDoesNotPublishItsResult() {
        var cache = new PdfReadCache<bool>();
        using var cancellation = new CancellationTokenSource();
        Assert.ThrowsAny<OperationCanceledException>(() => cache.GetOrCreate(cancellation, static (source, _) => {
            source.Cancel();
            return false;
        }, cancellation.Token));
        Assert.True(cache.GetOrCreate(true, static (value, _) => value));
        Assert.True(cache.GetOrCreate(false, static (value, _) => value));
    }

    private static async Task VerifyCancelledWait<T>(T expected) {
        var cache = new PdfReadCache<T>();
        using var entered = new ManualResetEventSlim();
        using var release = new ManualResetEventSlim();
        int initializations = 0;
        T Initialize(CancellationToken token) {
            Interlocked.Increment(ref initializations);
            entered.Set();
            if (!release.Wait(TimeSpan.FromSeconds(10), token)) throw new TimeoutException("Read initialization was not released.");
            return expected;
        }
        Func<CancellationToken, T> initialize = Initialize;
        Task<T> first = Task.Factory.StartNew(() => cache.GetOrCreate(initialize, static (factory, token) => factory(token)),
            CancellationToken.None, TaskCreationOptions.LongRunning, TaskScheduler.Default);
        Task<T>? waiting = null;
        try {
            Assert.True(entered.Wait(TimeSpan.FromSeconds(5)));
            using var cancellation = new CancellationTokenSource(TimeSpan.FromMilliseconds(100));
            waiting = Task.Factory.StartNew(() => cache.GetOrCreate(initialize, static (factory, token) => factory(token), cancellation.Token),
                CancellationToken.None, TaskCreationOptions.LongRunning, TaskScheduler.Default);
            Assert.Same(waiting, await Task.WhenAny(waiting, Task.Delay(TimeSpan.FromSeconds(2))));
            await Assert.ThrowsAnyAsync<OperationCanceledException>(async () => await waiting);
            Assert.False(first.IsCompleted);
        } finally {
            release.Set();
            await first;
            if (waiting is not null) { try { await waiting; } catch (OperationCanceledException) { } }
        }
        Assert.Equal(expected, await first);
        Assert.Equal(expected, cache.GetOrCreate(initialize, static (factory, token) => factory(token)));
        Assert.Equal(1, initializations);
    }
}
