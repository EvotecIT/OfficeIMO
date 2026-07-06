namespace OfficeIMO.Pdf;

/// <summary>
/// JSON-friendly summary of a PDF rewrite-preservation matrix report.
/// </summary>
public sealed class PdfRewritePreservationMatrixSummary {
    internal PdfRewritePreservationMatrixSummary(PdfRewritePreservationMatrixReport report) {
        Passed = report.Passed;
        Summary = report.Summary;
        Rows = report.Entries.Select(static entry => new PdfRewritePreservationMatrixRowSummary(entry)).ToArray();
    }

    /// <summary>True when every matrix row produced its expected classification.</summary>
    public bool Passed { get; }

    /// <summary>Human-readable matrix summary.</summary>
    public string Summary { get; }

    /// <summary>Scenario rows in execution order.</summary>
    public IReadOnlyList<PdfRewritePreservationMatrixRowSummary> Rows { get; }
}
