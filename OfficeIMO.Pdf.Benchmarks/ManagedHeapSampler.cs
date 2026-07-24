using System.Threading;

internal sealed class ManagedHeapSampler : IDisposable {
    private readonly Thread _thread;
    private readonly ManualResetEventSlim _stop = new(false);
    private long _peakBytes;
    private int _stopped;

    internal ManagedHeapSampler() {
        _peakBytes = GC.GetTotalMemory(forceFullCollection: false);
        _thread = new Thread(SampleUntilStopped) {
            IsBackground = true,
            Name = "OfficeIMO.Pdf managed heap sampler"
        };
        _thread.Start();
    }

    internal long Stop() {
        if (Interlocked.Exchange(ref _stopped, 1) == 0) {
            _stop.Set();
            _thread.Join();
            RecordCurrentHeap();
        }

        return Interlocked.Read(ref _peakBytes);
    }

    public void Dispose() {
        Stop();
        _stop.Dispose();
    }

    private void SampleUntilStopped() {
        while (!_stop.Wait(1)) {
            RecordCurrentHeap();
        }
    }

    private void RecordCurrentHeap() {
        long observed = GC.GetTotalMemory(forceFullCollection: false);
        long current = Interlocked.Read(ref _peakBytes);
        while (observed > current) {
            long prior = Interlocked.CompareExchange(ref _peakBytes, observed, current);
            if (prior == current) {
                return;
            }

            current = prior;
        }
    }
}
