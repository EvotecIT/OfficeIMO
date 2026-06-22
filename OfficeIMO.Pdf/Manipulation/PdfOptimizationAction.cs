namespace OfficeIMO.Pdf;

/// <summary>Lossless optimization action applied to a PDF object.</summary>
public sealed class PdfOptimizationAction {
    internal PdfOptimizationAction(
        string kind,
        int objectNumber,
        long originalLengthBytes,
        long optimizedLengthBytes,
        string description) {
        Kind = kind;
        ObjectNumber = objectNumber;
        OriginalLengthBytes = originalLengthBytes;
        OptimizedLengthBytes = optimizedLengthBytes;
        Description = description;
    }

    /// <summary>Stable action kind, for example CompressStream.</summary>
    public string Kind { get; }

    /// <summary>PDF object number affected by the action.</summary>
    public int ObjectNumber { get; }

    /// <summary>Original object payload size in bytes.</summary>
    public long OriginalLengthBytes { get; }

    /// <summary>Optimized object payload size in bytes.</summary>
    public long OptimizedLengthBytes { get; }

    /// <summary>Bytes saved by this action.</summary>
    public long SavedBytes => Math.Max(0, OriginalLengthBytes - OptimizedLengthBytes);

    /// <summary>Human-readable action description.</summary>
    public string Description { get; }
}
