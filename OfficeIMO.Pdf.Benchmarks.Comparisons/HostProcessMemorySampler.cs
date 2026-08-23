using System.Diagnostics;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

internal sealed class HostProcessMemorySampler : IAsyncDisposable {
    private readonly CancellationTokenSource _stop = new();
    private readonly Task _samplingTask;
    private long _peakBytes;

    internal HostProcessMemorySampler() {
        BeforeBytes = ReadWorkingSet();
        _peakBytes = BeforeBytes;
        _samplingTask = Task.Run(SampleAsync);
    }

    internal long BeforeBytes { get; }

    internal long AfterBytes { get; private set; }

    internal long PeakBytes => Math.Max(_peakBytes, Math.Max(BeforeBytes, AfterBytes));

    public async ValueTask DisposeAsync() {
        AfterBytes = ReadWorkingSet();
        _stop.Cancel();
        try {
            await _samplingTask.ConfigureAwait(false);
        } catch (OperationCanceledException) {
            // Normal sampler shutdown.
        } finally {
            _stop.Dispose();
        }
    }

    private async Task SampleAsync() {
        while (!_stop.IsCancellationRequested) {
            long current = ReadWorkingSet();
            long observed = Volatile.Read(ref _peakBytes);
            while (current > observed && Interlocked.CompareExchange(ref _peakBytes, current, observed) != observed) {
                observed = Volatile.Read(ref _peakBytes);
            }

            await Task.Delay(TimeSpan.FromMilliseconds(10), _stop.Token).ConfigureAwait(false);
        }
    }

    private static long ReadWorkingSet() {
        using Process process = Process.GetCurrentProcess();
        process.Refresh();
        return process.WorkingSet64;
    }
}
