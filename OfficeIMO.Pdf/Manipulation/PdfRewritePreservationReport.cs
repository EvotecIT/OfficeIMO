namespace OfficeIMO.Pdf;

/// <summary>
/// Result of comparing original and rewritten PDFs for user-visible preservation.
/// </summary>
public sealed class PdfRewritePreservationReport {
    private readonly PdfDocumentInfo _original;
    private readonly PdfDocumentInfo _rewritten;
    private readonly PdfPermissionPolicy _originalPolicy;
    private readonly PdfPermissionPolicy _rewrittenPolicy;

    internal PdfRewritePreservationReport(PdfDocumentInfo original, PdfDocumentInfo rewritten, IReadOnlyList<PdfRewritePreservationIssue> issues,
        PdfPermissionPolicy originalPolicy, PdfPermissionPolicy rewrittenPolicy) {
        _original = original; _rewritten = rewritten;
        _originalPolicy = originalPolicy; _rewrittenPolicy = rewrittenPolicy;
        Issues = issues;
    }

    /// <summary>Whether both complete inspection snapshots may be read with the supplied authorizations.</summary>
    public bool CanReadDetails => PdfPermissionAuthorization.CanExtractContent(_original.Security, _originalPolicy) &&
        PdfPermissionAuthorization.CanExtractContent(_rewritten.Security, _rewrittenPolicy);

    /// <summary>Inspection snapshot for the original PDF, subject to its content-extraction permission.</summary>
    /// <exception cref="PdfPermissionDeniedException">The original authorization does not permit content extraction.</exception>
    public PdfDocumentInfo Original {
        get { PdfPermissionAuthorization.DemandContentExtraction(_original.Security, _originalPolicy, "original rewrite-preservation report"); return _original; }
    }

    /// <summary>Inspection snapshot for the rewritten PDF, subject to its content-extraction permission.</summary>
    /// <exception cref="PdfPermissionDeniedException">The rewritten authorization does not permit content extraction.</exception>
    public PdfDocumentInfo Rewritten {
        get { PdfPermissionAuthorization.DemandContentExtraction(_rewritten.Security, _rewrittenPolicy, "rewritten rewrite-preservation report"); return _rewritten; }
    }

    /// <summary>Preservation mismatches found in the rewritten PDF.</summary>
    public IReadOnlyList<PdfRewritePreservationIssue> Issues { get; }

    /// <summary>True when no preservation mismatches were found.</summary>
    public bool IsPreserved => Issues.Count == 0;

    /// <summary>Human-readable summary suitable for logs, tests, and wrappers.</summary>
    public string Summary {
        get {
            if (IsPreserved) {
                return "PDF rewrite preservation checks passed.";
            }

            return "PDF rewrite preservation failed: " + string.Join("; ", Issues.Select(issue => issue.Message));
        }
    }

    /// <summary>Throws an InvalidOperationException when preservation checks found mismatches.</summary>
    public void ThrowIfFailed() {
        if (!IsPreserved) {
            throw new InvalidOperationException(Summary);
        }
    }
}
