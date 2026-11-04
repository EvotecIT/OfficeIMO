using System.Diagnostics;
using System.Text.Json;

namespace OfficeIMO.Pdf.Benchmarks;

public sealed partial class PdfRuntimeWorkload {
    private readonly List<PhaseMemoryBoundary> _phaseMemory = new(8);
    private FirstWriteFileStream? _firstWrite;
    private long _phaseOperationStart;
    private long _phaseSaveStart;

    /// <summary>Time from benchmark method entry before composition to first destination-write completion.</summary>
    public double OperationToFirstWriteCompletionMs => EventMilliseconds(_phaseOperationStart, RequireFirstWrite().FirstWriteCompletionTimestamp);
    /// <summary>Time from immediately before Save to first destination-write completion; includes the copy for the copied strategy.</summary>
    public double SaveToFirstWriteCompletionMs => EventMilliseconds(_phaseSaveStart, RequireFirstWrite().FirstWriteCompletionTimestamp);
    /// <summary>Time from benchmark method entry to the first successful nonempty underlying write call.</summary>
    public double OperationToFirstWriteCallStartMs => EventMilliseconds(_phaseOperationStart, RequireFirstWrite().FirstWriteCallStartTimestamp);
    /// <summary>Time from immediately before Save to the first successful nonempty underlying write call.</summary>
    public double SaveToFirstWriteCallStartMs => EventMilliseconds(_phaseSaveStart, RequireFirstWrite().FirstWriteCallStartTimestamp);
    /// <summary>Elapsed interval of that first underlying FileStream write call.</summary>
    public double FirstWriteCallIntervalMs => EventMilliseconds(RequireFirstWrite().FirstWriteCallStartTimestamp, RequireFirstWrite().FirstWriteCompletionTimestamp);
    /// <summary>Payload of the first successful nonempty underlying FileStream write.</summary>
    public int FirstWritePayloadBytes => RequireFirstWrite().FirstWritePayloadBytes;
    /// <summary>Bytes accepted across all successful nonempty destination writes.</summary>
    public long DestinationWrittenBytes => RequireFirstWrite().TotalWrittenBytes;
    /// <summary>Successful nonempty write count at the observed destination.</summary>
    public long DestinationNonemptyWriteCalls => RequireFirstWrite().NonemptyWriteCalls;

    /// <summary>Releases previous operation and diagnostic observations before the shared runner measures an iteration.</summary>
    public void PrepareFirstWrite() {
        ReleaseResults();
        _phaseMemory.Clear();
        _firstWrite = null;
        _phaseOperationStart = 0L;
        _phaseSaveStart = 0L;
    }

    /// <summary>
    /// Composes and saves a complete file with benchmark-only first-write and noncollecting memory observations.
    /// The shared PowerForge runner owns total elapsed time, allocation and periodic memory sampling.
    /// </summary>
    public void ExecuteObservedDestination(string operation, string destination, CancellationToken cancellationToken = default) {
        if (operation is not ("DestinationDirect" or "DestinationCopied")) throw new ArgumentException("Unknown destination operation.", nameof(operation));
        if (_expectedDestinationSha256 is null) throw new InvalidOperationException("Destination reference was not prepared.");
        cancellationToken.ThrowIfCancellationRequested();
        _phaseOperationStart = Stopwatch.GetTimestamp();
        _operation = operation;
        _destinationPath = destination;
        ObserveMemory("OperationEntry");
        PdfDocument document = PdfRuntimeCorpus.Create(Pages, PdfObjectSerializationMode.ForwardOnly);
        ObserveMemory("CompositionComplete");
        PdfSaveResult saved;
        if (operation == "DestinationDirect") {
            using (var stream = new FirstWriteFileStream(destination, () => ObserveMemory("FirstWriteComplete"))) {
                _firstWrite = stream;
                ObserveMemory("BeforeSave");
                _phaseSaveStart = Stopwatch.GetTimestamp();
                saved = document.SaveWithEvidence(stream, cancellationToken);
                ObserveMemory("SaveComplete");
            }
        } else {
            using var memory = new MemoryStream();
            ObserveMemory("BeforeSave");
            _phaseSaveStart = Stopwatch.GetTimestamp();
            saved = document.SaveWithEvidence(memory, cancellationToken);
            ObserveMemory("SaveComplete");
            byte[] bytes = memory.ToArray();
            cancellationToken.ThrowIfCancellationRequested();
            ObserveMemory("BeforeDestinationWrite");
            using (var stream = new FirstWriteFileStream(destination, () => ObserveMemory("FirstWriteComplete"))) {
                _firstWrite = stream;
                stream.Write(bytes);
            }
        }
        _serialization = saved.Serialization ?? throw new InvalidDataException("Generation returned no serialization report.");
        ObserveMemory("DestinationClosed");
    }

    /// <summary>Returns a named boundary metric after the measured operation without forcing collection.</summary>
    public long PhaseMemoryBytes(string boundary, string kind) {
        PhaseMemoryBoundary snapshot = _phaseMemory.Single(item => item.Name == boundary);
        return kind switch {
            "ManagedHeap" => snapshot.ManagedHeapBytes,
            "WorkingSet" => snapshot.WorkingSetBytes,
            _ => throw new ArgumentException("Unknown memory observation kind.", nameof(kind))
        };
    }

    /// <summary>Qualifies the persisted destination and precise event ordering outside the measured operation.</summary>
    public void ValidateFirstWrite(string directory, double powerForgeDurationMs) {
        ValidateDestination(directory);
        FirstWriteFileStream first = RequireFirstWrite();
        if (_phaseOperationStart <= 0L || _phaseSaveStart < _phaseOperationStart ||
            first.FirstWriteCallStartTimestamp < _phaseSaveStart ||
            first.FirstWriteCompletionTimestamp < first.FirstWriteCallStartTimestamp || first.FirstWritePayloadBytes <= 0 ||
            first.TotalWrittenBytes != OutputBytes || first.FirstWritePayloadBytes > OutputBytes ||
            OperationToFirstWriteCompletionMs > powerForgeDurationMs + 1D || _phaseMemory.Count != (_operation == "DestinationCopied" ? 7 : 6)) {
            throw new InvalidDataException("First-write event ordering, byte count or memory boundaries changed.");
        }
        if (_operation == "DestinationCopied" && (first.FirstWritePayloadBytes != OutputBytes || first.NonemptyWriteCalls != 1L)) {
            throw new InvalidDataException("Copied output did not write its complete artifact through the same observed destination.");
        }
        File.WriteAllText(Path.Combine(directory, "first-write-evidence.json"), JsonSerializer.Serialize(new {
            Contract = "pdf-first-successful-destination-write-v1", Operation = _operation, CorpusProfile, Pages, SourceSha256,
            ExpectedDestinationSha256, DestinationSha256, InputBytes, OutputBytes, VerifiedPages, FinalByteArrayMaterialized,
            EventClock = "Stopwatch.GetTimestamp, benchmark-local monotonic event observations; PowerForge alone measures total operation duration",
            EventClockFrequency = Stopwatch.Frequency, OperationStartTimestamp = _phaseOperationStart, SaveStartTimestamp = _phaseSaveStart,
            first.FirstWriteCallStartTimestamp, first.FirstWriteCompletionTimestamp, OperationToFirstWriteCallStartMs,
            OperationToFirstWriteCompletionMs, SaveToFirstWriteCallStartMs, SaveToFirstWriteCompletionMs, FirstWriteCallIntervalMs,
            FirstWritePayloadBytes, DestinationWrittenBytes, DestinationNonemptyWriteCalls, PowerForgeDurationMs = powerForgeDurationMs,
            Boundaries = _phaseMemory,
            Qualification = "First successful nonempty destination-write completion is an observable upper-bound proxy for acceptance by FileStream/OS cache, not first physical byte, durable persistence or network delivery. Named memory observations do not force GC, may miss peaks and include observer/host effects. The copied lane additionally observes BeforeDestinationWrite after ToArray; direct writes happen within Save. Separate instrumented cohort.",
            Serialization = _serialization
        }, new JsonSerializerOptions { WriteIndented = true }));
    }

    private void ObserveMemory(string name) {
        long managed = GC.GetTotalMemory(forceFullCollection: false);
        using Process process = Process.GetCurrentProcess();
        process.Refresh();
        _phaseMemory.Add(new PhaseMemoryBoundary(name, managed, process.WorkingSet64));
    }

    private FirstWriteFileStream RequireFirstWrite() => _firstWrite ?? throw new InvalidOperationException("No first-write observation was produced.");
    private static double EventMilliseconds(long start, long end) => (end - start) * 1000D / Stopwatch.Frequency;
    private sealed record PhaseMemoryBoundary(string Name, long ManagedHeapBytes, long WorkingSetBytes);
}
