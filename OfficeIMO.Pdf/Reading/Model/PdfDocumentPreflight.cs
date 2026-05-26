namespace OfficeIMO.Pdf;

/// <summary>
/// Wrapper-friendly PDF capability report for OfficeIMO.Pdf read and rewrite operations.
/// </summary>
public sealed class PdfDocumentPreflight {
    internal PdfDocumentPreflight(
        PdfDocumentProbe probe,
        PdfDocumentInfo? documentInfo,
        bool canRead,
        bool canRewrite,
        IReadOnlyList<string> diagnostics,
        IReadOnlyList<PdfReadBlocker> readBlockers,
        IReadOnlyList<PdfRewriteBlocker> rewriteBlockers) {
        Probe = probe;
        DocumentInfo = documentInfo;
        CanRead = canRead;
        CanRewrite = canRewrite;
        Diagnostics = diagnostics;
        ReadBlockers = readBlockers;
        RewriteBlockers = rewriteBlockers;
    }

    /// <summary>Lightweight PDF markers read before full parsing.</summary>
    public PdfDocumentProbe Probe { get; }

    /// <summary>Parsed document information when the document can be inspected.</summary>
    public PdfDocumentInfo? DocumentInfo { get; }

    /// <summary>True when OfficeIMO.Pdf can parse enough of the document for read-oriented operations.</summary>
    public bool CanRead { get; }

    /// <summary>True when OfficeIMO.Pdf can attempt rewrite-style manipulation without known security blockers.</summary>
    public bool CanRewrite { get; }

    /// <summary>Human-readable diagnostics explaining blocked or risky operations.</summary>
    public IReadOnlyList<string> Diagnostics { get; }

    /// <summary>Structured reasons why read-oriented operations are blocked.</summary>
    public IReadOnlyList<PdfReadBlocker> ReadBlockers { get; }

    /// <summary>Structured reasons why rewrite-style manipulation is blocked.</summary>
    public IReadOnlyList<PdfRewriteBlocker> RewriteBlockers { get; }

    /// <summary>Returns true when a specific read blocker is present.</summary>
    public bool HasReadBlocker(PdfReadBlockerKind kind) {
        for (int i = 0; i < ReadBlockers.Count; i++) {
            if (ReadBlockers[i].Kind == kind) {
                return true;
            }
        }

        return false;
    }

    /// <summary>Returns true when a specific rewrite blocker is present.</summary>
    public bool HasRewriteBlocker(PdfRewriteBlockerKind kind) {
        for (int i = 0; i < RewriteBlockers.Count; i++) {
            if (RewriteBlockers[i].Kind == kind) {
                return true;
            }
        }

        return false;
    }
}
