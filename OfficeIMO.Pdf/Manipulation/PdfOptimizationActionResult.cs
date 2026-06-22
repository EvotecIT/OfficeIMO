namespace OfficeIMO.Pdf;

/// <summary>Result of applying lossless PDF optimization actions.</summary>
public sealed class PdfOptimizationActionResult {
    internal PdfOptimizationActionResult(
        byte[] bytes,
        long originalLengthBytes,
        long optimizedLengthBytes,
        PdfOptimizationReport reportBefore,
        IReadOnlyList<PdfOptimizationAction> actions,
        bool returnedOriginal) {
        Bytes = bytes;
        OriginalLengthBytes = originalLengthBytes;
        OptimizedLengthBytes = optimizedLengthBytes;
        ReportBefore = reportBefore;
        Actions = actions;
        ReturnedOriginal = returnedOriginal;
    }

    /// <summary>Optimized PDF bytes, or original bytes when no smaller optimized form was produced.</summary>
    public byte[] Bytes { get; }

    /// <summary>Input PDF length in bytes.</summary>
    public long OriginalLengthBytes { get; }

    /// <summary>Returned PDF length in bytes.</summary>
    public long OptimizedLengthBytes { get; }

    /// <summary>Bytes saved in the returned PDF.</summary>
    public long SavedBytes => Math.Max(0, OriginalLengthBytes - OptimizedLengthBytes);

    /// <summary>Optimization analysis collected before applying actions.</summary>
    public PdfOptimizationReport ReportBefore { get; }

    /// <summary>Actions applied while building the optimized candidate.</summary>
    public IReadOnlyList<PdfOptimizationAction> Actions { get; }

    /// <summary>True when the original bytes were returned because the optimized candidate was not smaller.</summary>
    public bool ReturnedOriginal { get; }

    /// <summary>True when the returned bytes are smaller than the input bytes.</summary>
    public bool Applied => !ReturnedOriginal && SavedBytes > 0;

    /// <summary>Number of actions applied while building the optimized candidate.</summary>
    public int ActionCount => Actions.Count;
}
