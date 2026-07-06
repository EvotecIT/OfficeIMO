namespace OfficeIMO.Pdf;

/// <summary>
/// JSON-friendly summary of one PDF rewrite-preservation issue.
/// </summary>
public sealed class PdfRewritePreservationIssueSummary {
    internal PdfRewritePreservationIssueSummary(PdfRewritePreservationIssue issue) {
        Feature = issue.Feature;
        Expected = issue.Expected;
        Actual = issue.Actual;
        Message = issue.Message;
    }

    /// <summary>Stable feature name.</summary>
    public string Feature { get; }

    /// <summary>Expected value.</summary>
    public string Expected { get; }

    /// <summary>Observed value.</summary>
    public string Actual { get; }

    /// <summary>Human-readable preservation diagnostic.</summary>
    public string Message { get; }
}
