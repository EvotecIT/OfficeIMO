namespace OfficeIMO.Pdf;

/// <summary>
/// Machine-readable reason why read-oriented PDF operations are blocked.
/// </summary>
public sealed class PdfReadBlocker {
    internal PdfReadBlocker(PdfReadBlockerKind kind, string message) {
        Kind = kind;
        Message = message;
    }

    /// <summary>Category of read blocker.</summary>
    public PdfReadBlockerKind Kind { get; }

    /// <summary>Human-readable diagnostic for logs and wrapper errors.</summary>
    public string Message { get; }
}
