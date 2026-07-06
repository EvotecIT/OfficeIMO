namespace OfficeIMO.Pdf;

/// <summary>
/// Result of running a PDF rewrite-preservation proof matrix.
/// </summary>
public sealed class PdfRewritePreservationMatrixReport {
    internal PdfRewritePreservationMatrixReport(IReadOnlyList<PdfRewritePreservationMatrixEntry> entries) {
        Entries = entries;
    }

    /// <summary>Scenario results in input order.</summary>
    public IReadOnlyList<PdfRewritePreservationMatrixEntry> Entries { get; }

    /// <summary>True when every scenario produced the expected classification.</summary>
    public bool Passed => Entries.All(static entry => entry.Passed);

    /// <summary>Human-readable summary for logs and proof files.</summary>
    public string Summary {
        get {
            if (Passed) {
                return "PDF rewrite preservation matrix passed for " + Entries.Count.ToString(System.Globalization.CultureInfo.InvariantCulture) + " scenario(s).";
            }

            return "PDF rewrite preservation matrix failed: " + string.Join("; ", Entries.Where(static entry => !entry.Passed).Select(static entry => entry.Summary));
        }
    }

    /// <summary>Throws an InvalidOperationException when any scenario produced an unexpected classification.</summary>
    public void ThrowIfFailed() {
        if (!Passed) {
            throw new InvalidOperationException(Summary);
        }
    }

    /// <summary>Creates a JSON-friendly summary of this matrix report for proof packs, CI logs, and wrappers.</summary>
    public PdfRewritePreservationMatrixSummary ToSummary() {
        return new PdfRewritePreservationMatrixSummary(this);
    }
}
