namespace OfficeIMO.Pdf;

/// <summary>Typed preview of items that the supplied PDF sanitization policy would remove.</summary>
public sealed class PdfSanitizationReport {
    internal PdfSanitizationReport(IReadOnlyList<PdfSanitizationFinding> findings) {
        Findings = findings;
        ActionCounts = new PdfSanitizationActionCounts(findings);
    }

    /// <summary>Items selected for removal by the inspected policy.</summary>
    public IReadOnlyList<PdfSanitizationFinding> Findings { get; }

    /// <summary>Per-kind action counts for the selected policy.</summary>
    public PdfSanitizationActionCounts ActionCounts { get; }

    /// <summary>Total number of selected findings.</summary>
    public int TotalCount => Findings.Count;
}
