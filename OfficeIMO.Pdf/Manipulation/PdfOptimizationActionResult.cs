namespace OfficeIMO.Pdf;

/// <summary>Result of applying lossless PDF optimization actions.</summary>
public sealed class PdfOptimizationActionResult {
    private readonly byte[] _bytes;
    private readonly PdfReadOptions? _readOptions;

    internal PdfOptimizationActionResult(
        byte[] bytes,
        long originalLengthBytes,
        long optimizedLengthBytes,
        long candidateLengthBytes,
        PdfOptimizationReport reportBefore,
        PdfOptimizationReport reportAfter,
        IReadOnlyList<PdfOptimizationAction> actions,
        IReadOnlyList<PdfOptimizationSkippedAction> skippedActions,
        PdfRewritePreservationReport preservationReport,
        PdfOptimizationProfile requestedProfile,
        PdfOptimizationXrefFormat candidateXrefFormat,
        bool candidateUsesObjectStreams,
        bool candidateLinearized,
        bool returnedOriginal,
        PdfReadOptions? readOptions = null) {
        _bytes = (byte[])bytes.Clone();
        _readOptions = PdfReadOptions.WithMinimumInputBytes(readOptions, _bytes.LongLength);
        OriginalLengthBytes = originalLengthBytes;
        OptimizedLengthBytes = optimizedLengthBytes;
        CandidateLengthBytes = candidateLengthBytes;
        ReportBefore = reportBefore;
        ReportAfter = reportAfter;
        Actions = actions.ToArray();
        SkippedActions = skippedActions.ToArray();
        PreservationReport = preservationReport;
        RequestedProfile = requestedProfile;
        CandidateXrefFormat = candidateXrefFormat;
        CandidateUsesObjectStreams = candidateUsesObjectStreams;
        CandidateLinearized = candidateLinearized;
        ReturnedOriginal = returnedOriginal;
    }

    /// <summary>Optimized PDF bytes, or original bytes when no smaller optimized form was produced.</summary>
    public byte[] Bytes => (byte[])_bytes.Clone();

    /// <summary>Input PDF length in bytes.</summary>
    public long OriginalLengthBytes { get; }

    /// <summary>Returned PDF length in bytes.</summary>
    public long OptimizedLengthBytes { get; }

    /// <summary>Optimized candidate length before KeepOriginalWhenNotSmaller is applied.</summary>
    public long CandidateLengthBytes { get; }

    /// <summary>Bytes saved in the returned PDF.</summary>
    public long SavedBytes => Math.Max(0, OriginalLengthBytes - OptimizedLengthBytes);

    /// <summary>Bytes the optimized candidate would save before KeepOriginalWhenNotSmaller is applied.</summary>
    public long CandidateSavedBytes => Math.Max(0, OriginalLengthBytes - CandidateLengthBytes);

    /// <summary>Optimization analysis collected before applying actions.</summary>
    public PdfOptimizationReport ReportBefore { get; }

    /// <summary>Optimization analysis collected from the optimized candidate.</summary>
    public PdfOptimizationReport ReportAfter { get; }

    /// <summary>Actions applied while building the optimized candidate.</summary>
    public IReadOnlyList<PdfOptimizationAction> Actions { get; }

    /// <summary>Optimization opportunities skipped while building the candidate.</summary>
    public IReadOnlyList<PdfOptimizationSkippedAction> SkippedActions { get; }

    /// <summary>Post-save proof that non-size document contracts were preserved.</summary>
    public PdfRewritePreservationReport PreservationReport { get; }

    /// <summary>Named profile requested for the candidate.</summary>
    public PdfOptimizationProfile RequestedProfile { get; }

    /// <summary>Cross-reference representation emitted by the candidate.</summary>
    public PdfOptimizationXrefFormat CandidateXrefFormat { get; }

    /// <summary>True when the candidate packs eligible objects into object streams.</summary>
    public bool CandidateUsesObjectStreams { get; }

    /// <summary>True when the candidate uses standards-compliant Fast Web View layout and hint tables.</summary>
    public bool CandidateLinearized { get; }

    /// <summary>True when the original bytes were returned because the optimized candidate was not smaller.</summary>
    public bool ReturnedOriginal { get; }

    /// <summary>True when the returned bytes are smaller than the input bytes.</summary>
    public bool Applied => !ReturnedOriginal && SavedBytes > 0;

    /// <summary>Number of actions applied while building the optimized candidate.</summary>
    public int ActionCount => Actions.Count;

    /// <summary>Number of skipped optimization opportunities recorded while building the candidate.</summary>
    public int SkippedActionCount => SkippedActions.Count;

    /// <summary>Opens the selected optimized or original bytes through the fluent document API.</summary>
    public PdfDocument ToDocument(PdfReadOptions? readOptions = null) => PdfDocument.Open(_bytes, readOptions ?? _readOptions);

}
