using System.Diagnostics;

namespace OfficeIMO.Html.Benchmarks;

internal sealed class HtmlLayoutMemorySampler : IDisposable {
    private readonly Process _process;
    private readonly Thread _thread;
    private readonly ManualResetEventSlim _stop = new(false);
    private long _peakManaged;
    private long _peakWorkingSet;
    private int _stopped;

    internal HtmlLayoutMemorySampler(Process process) {
        _process = process;
        _peakManaged = GC.GetTotalMemory(false);
        process.Refresh();
        _peakWorkingSet = process.WorkingSet64;
        _thread = new Thread(Sample) { IsBackground = true, Name = "OfficeIMO.Html layout memory sampler" };
        _thread.Start();
    }

    internal HtmlLayoutMemoryPeak Stop() {
        if (Interlocked.Exchange(ref _stopped, 1) == 0) { _stop.Set(); _thread.Join(); Record(); }
        return new HtmlLayoutMemoryPeak(Interlocked.Read(ref _peakManaged), Interlocked.Read(ref _peakWorkingSet));
    }

    public void Dispose() { Stop(); _stop.Dispose(); }
    private void Sample() { while (!_stop.Wait(1)) Record(); }
    private void Record() { RecordPeak(ref _peakManaged, GC.GetTotalMemory(false)); _process.Refresh(); RecordPeak(ref _peakWorkingSet, _process.WorkingSet64); }
    private static void RecordPeak(ref long target, long observed) { long current = Interlocked.Read(ref target); while (observed > current) { long prior = Interlocked.CompareExchange(ref target, observed, current); if (prior == current) return; current = prior; } }
}
