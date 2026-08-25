using System.Diagnostics;

namespace OfficeIMO.Email.Benchmarks.Comparisons;

internal sealed class EmailMimeMemorySampler : IDisposable {
    private readonly Process _process;
    private readonly Thread _thread;
    private readonly ManualResetEventSlim _stop = new(false);
    private long _peakManagedHeapBytes;
    private long _peakWorkingSetBytes;
    private int _stopped;

    internal EmailMimeMemorySampler(Process process) {
        _process = process;
        _peakManagedHeapBytes = GC.GetTotalMemory(forceFullCollection: false);
        _process.Refresh();
        _peakWorkingSetBytes = _process.WorkingSet64;
        _thread = new Thread(SampleUntilStopped) {
            IsBackground = true,
            Name = "OfficeIMO.Email MIME memory sampler"
        };
        _thread.Start();
    }

    internal EmailMimeMemoryPeak Stop() {
        if (Interlocked.Exchange(ref _stopped, 1) == 0) {
            _stop.Set();
            _thread.Join();
            RecordCurrentMemory();
        }
        return new EmailMimeMemoryPeak(
            Interlocked.Read(ref _peakManagedHeapBytes),
            Interlocked.Read(ref _peakWorkingSetBytes));
    }

    public void Dispose() {
        Stop();
        _stop.Dispose();
    }

    private void SampleUntilStopped() {
        while (!_stop.Wait(1)) RecordCurrentMemory();
    }

    private void RecordCurrentMemory() {
        RecordPeak(ref _peakManagedHeapBytes, GC.GetTotalMemory(forceFullCollection: false));
        _process.Refresh();
        RecordPeak(ref _peakWorkingSetBytes, _process.WorkingSet64);
    }

    private static void RecordPeak(ref long peak, long observed) {
        long current = Interlocked.Read(ref peak);
        while (observed > current) {
            long prior = Interlocked.CompareExchange(ref peak, observed, current);
            if (prior == current) return;
            current = prior;
        }
    }
}
