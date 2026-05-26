namespace OfficeIMO.Pdf;

/// <summary>
/// Machine-readable reason why rewrite-style PDF manipulation is blocked.
/// </summary>
public sealed class PdfRewriteBlocker {
    internal PdfRewriteBlocker(PdfRewriteBlockerKind kind, string message) {
        Kind = kind;
        Message = message;
    }

    /// <summary>Category of rewrite blocker.</summary>
    public PdfRewriteBlockerKind Kind { get; }

    /// <summary>Human-readable diagnostic for logs and wrapper errors.</summary>
    public string Message { get; }
}
