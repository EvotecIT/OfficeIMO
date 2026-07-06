namespace OfficeIMO.Pdf;

/// <summary>
/// Result of comparing original and rewritten PDFs for user-visible preservation.
/// </summary>
public sealed class PdfRewritePreservationReport {
    internal PdfRewritePreservationReport(PdfDocumentInfo original, PdfDocumentInfo rewritten, IReadOnlyList<PdfRewritePreservationIssue> issues) {
        Original = original;
        Rewritten = rewritten;
        Issues = issues;
    }

    /// <summary>Inspection snapshot for the original PDF.</summary>
    public PdfDocumentInfo Original { get; }

    /// <summary>Inspection snapshot for the rewritten PDF.</summary>
    public PdfDocumentInfo Rewritten { get; }

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
