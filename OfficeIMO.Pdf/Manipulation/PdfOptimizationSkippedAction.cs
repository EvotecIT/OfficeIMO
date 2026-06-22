namespace OfficeIMO.Pdf;

/// <summary>Lossless optimization opportunity that was not applied.</summary>
public sealed class PdfOptimizationSkippedAction {
    internal PdfOptimizationSkippedAction(
        string kind,
        int objectNumber,
        long originalLengthBytes,
        string reason,
        string description) {
        Kind = kind;
        ObjectNumber = objectNumber;
        OriginalLengthBytes = originalLengthBytes;
        Reason = reason;
        Description = description;
    }

    /// <summary>Stable action kind, for example CompressStream.</summary>
    public string Kind { get; }

    /// <summary>PDF object number considered for the action.</summary>
    public int ObjectNumber { get; }

    /// <summary>Original object payload size in bytes.</summary>
    public long OriginalLengthBytes { get; }

    /// <summary>Stable reason code explaining why the action was skipped.</summary>
    public string Reason { get; }

    /// <summary>Human-readable skipped-action description.</summary>
    public string Description { get; }
}
