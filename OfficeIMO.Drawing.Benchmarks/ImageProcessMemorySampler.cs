using System.Diagnostics;

namespace OfficeIMO.Drawing.Benchmarks;

/// <summary>Samples process and managed-heap peaks across one image operation.</summary>
internal sealed class ImageProcessMemorySampler : IDisposable {
    private readonly ManualResetEventSlim _started = new(false);
    private readonly Thread _thread;
    private volatile bool _stop;
    private long _baselineWorkingSet;
    private long _baselinePrivateBytes;
    private long _baselineManagedHeap;
    private long _baselineNativeBytesEstimate;
    private long _peakWorkingSet;
    private long _peakPrivateBytes;
    private long _peakManagedHeap;
    private long _peakNativeBytesEstimate;

    internal ImageProcessMemorySampler() {
        _thread = new Thread(Sample) {
            IsBackground = true,
            Name = "OfficeIMO image memory sampler"
        };
    }

    internal long PeakWorkingSetDelta => Math.Max(0L, _peakWorkingSet - _baselineWorkingSet);
    internal long PeakPrivateBytesDelta => Math.Max(0L, _peakPrivateBytes - _baselinePrivateBytes);
    internal long PeakManagedHeapDelta => Math.Max(0L, _peakManagedHeap - _baselineManagedHeap);
    internal long PeakNativeBytesEstimate => Math.Max(0L, _peakNativeBytesEstimate - _baselineNativeBytesEstimate);

    internal void Start() {
        _thread.Start();
        _started.Wait();
    }

    internal void Stop() {
        if (_stop) return;
        _stop = true;
        _thread.Join();
    }

    public void Dispose() {
        if (_thread.IsAlive) Stop();
        _started.Dispose();
    }

    private void Sample() {
        using Process process = Process.GetCurrentProcess();
        process.Refresh();
        _baselineWorkingSet = process.WorkingSet64;
        _baselinePrivateBytes = process.PrivateMemorySize64;
        _baselineManagedHeap = GC.GetTotalMemory(forceFullCollection: false);
        _baselineNativeBytesEstimate = EstimateNativeBytes(_baselinePrivateBytes, _baselineManagedHeap);
        _peakWorkingSet = _baselineWorkingSet;
        _peakPrivateBytes = _baselinePrivateBytes;
        _peakManagedHeap = _baselineManagedHeap;
        _peakNativeBytesEstimate = _baselineNativeBytesEstimate;
        _started.Set();
        while (!_stop) {
            Record(process);
            Thread.Sleep(1);
        }
        Record(process);
    }

    private void Record(Process process) {
        process.Refresh();
        long workingSet = process.WorkingSet64;
        long privateBytes = process.PrivateMemorySize64;
        long managedHeap = GC.GetTotalMemory(forceFullCollection: false);
        _peakWorkingSet = Math.Max(_peakWorkingSet, workingSet);
        _peakPrivateBytes = Math.Max(_peakPrivateBytes, privateBytes);
        _peakManagedHeap = Math.Max(_peakManagedHeap, managedHeap);
        _peakNativeBytesEstimate = Math.Max(_peakNativeBytesEstimate, EstimateNativeBytes(privateBytes, managedHeap));
    }

    private static long EstimateNativeBytes(long privateBytes, long managedHeap) =>
        Math.Max(0L, privateBytes - managedHeap);
}
