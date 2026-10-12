using System.Diagnostics;

namespace OfficeIMO.IWork.Benchmarks;

public sealed partial class IWorkRuntimeWorkload {
    private IWorkSourceDocument? _phaseSource;
    private string? _phaseSourceHash;
    private bool _phaseDestinationUnchanged;

    /// <summary>Gets elapsed time before the timed phase cancellation request.</summary>
    public double CancellationPreRequestElapsedMs { get; private set; }
    /// <summary>Gets process CPU advancement before the phase cancellation request, including observer activity.</summary>
    public double CancellationPreRequestCpuMs { get; private set; }
    /// <summary>Gets every source unit verified by a subsequent uncancelled projection.</summary>
    public int CancellationRecoveryUnits { get; private set; }

    private static bool IsPhaseCancellation(string? operation) =>
        operation is "CancelDuringProjection" or "CancelDuringNativeEncoding";

    private void RequirePhaseCancellationProfile(string operation) {
        if (_kind != IWorkDocumentKind.Keynote
            || (_createdNativeDocument is null && !(_nativeKeynoteScale && operation == "CancelDuringProjection")))
            throw new InvalidOperationException("Phase cancellation requires owned native Keynote, or NativeLarge for projection only.");
    }

    private void PreparePhaseCancellation(string operation) {
        RequirePhaseCancellationProfile(operation);
        _phaseSource = IWorkSourceDocument.Open(_input);
        _phaseSourceHash = _phaseSource.ComputePackageContentHash();
        _phaseDestinationUnchanged = false;
    }

    private void ExecutePhaseCancellation(string operation) {
        RequirePhaseCancellationProfile(operation);
        var source = _phaseSource ?? throw new InvalidOperationException("Prepare the phase cancellation operation first.");
        using var cancellation = new CancellationTokenSource();
        var view = source.WithCancellation(cancellation.Token);
        using var destination = new MemoryStream();
        destination.WriteByte(0x5a);
        using var process = Process.GetCurrentProcess();
        TimeSpan cpuBefore = default;
        long started = 0, requested = 0;
        int finished = 0;
        var timer = new Timer(_ => {
            if (Volatile.Read(ref finished) != 0) return;
            process.Refresh();
            CancellationPreRequestCpuMs = (process.TotalProcessorTime - cpuBefore).TotalMilliseconds;
            long now = Stopwatch.GetTimestamp();
            CancellationPreRequestElapsedMs = Stopwatch.GetElapsedTime(started, now).TotalMilliseconds;
            Volatile.Write(ref requested, now);
            cancellation.Cancel();
        }, null, Timeout.Infinite, Timeout.Infinite);
        try {
            process.Refresh(); cpuBefore = process.TotalProcessorTime;
            started = Stopwatch.GetTimestamp();
            timer.Change(25, Timeout.Infinite);
            if (operation == "CancelDuringProjection") _projection = view.ReadKeynote();
            else _createdNativeDocument!.Save(destination, cancellationToken: cancellation.Token);
            throw new InvalidDataException("The phase operation completed before observing cancellation; use a larger workload.");
        } catch (OperationCanceledException) {
            long observed = Stopwatch.GetTimestamp();
            long request = Volatile.Read(ref requested);
            if (!cancellation.IsCancellationRequested || request == 0)
                throw new InvalidDataException("The phase did not observe its timed cancellation request.");
            CancellationObserved = true;
            CancellationLatencyMs = Stopwatch.GetElapsedTime(request, observed).TotalMilliseconds;
            // Native Save completes encoding before copying. An untouched stream locates this cancellation before copy.
            _phaseDestinationUnchanged = destination.CanWrite && destination.Position == 1
                && destination.Length == 1 && destination.GetBuffer()[0] == 0x5a;
        } finally {
            Volatile.Write(ref finished, 1);
            // Join any callback before releasing its token source and process handle.
            timer.DisposeAsync().AsTask().GetAwaiter().GetResult();
        }
    }

    private void ValidatePhaseCancellation() {
        if (!CancellationObserved || CancellationPreRequestElapsedMs < 20 || CancellationPreRequestCpuMs < 5
            || !_phaseDestinationUnchanged || _projection is not null || _output is not null || _phaseSource is null
            || _phaseSource.ComputePackageContentHash() != _phaseSourceHash)
            throw new InvalidDataException("Phase cancellation lacks active-work, immutable-source or untouched-destination evidence.");
        var recovered = _phaseSource.WithCancellation(CancellationToken.None).ReadKeynote();
        if (!recovered.HasEditableContent) throw new InvalidDataException("The source did not recover after cancellation.");
        VerifiedUnits = 0;
        if (_nativeKeynoteScale) ValidateNativeKeynote(recovered);
        else ValidateText(recovered.Slides.Select(slide => slide.TextBoxes.Single().Content.PlainText), paragraphTerminator: "\n");
        if (VerifiedUnits != _units) throw new InvalidDataException("The cancelled source did not recover every expected slide.");
        CancellationRecoveryUnits = VerifiedUnits;
        VerifiedUnits = 1; // One observed boundary; recovery units remain a separate complete-content count.
    }
}
