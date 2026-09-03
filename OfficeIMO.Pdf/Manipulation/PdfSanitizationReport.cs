namespace OfficeIMO.Pdf;

/// <summary>Typed preview of items that the supplied PDF sanitization policy would remove.</summary>
public sealed class PdfSanitizationReport {
    internal PdfSanitizationReport(
        IReadOnlyList<PdfSanitizationFinding> findings,
        int userMetadata = 0,
        int embeddedFiles = 0,
        int commentsAndMarkup = 0,
        int bookmarks = 0,
        int optionalContent = 0) {
        Findings = findings;
        ActionCounts = new PdfSanitizationActionCounts(findings);
        CategoryCounts = new PdfSanitizationCategoryCounts(
            userMetadata,
            embeddedFiles,
            ActionCounts.Total,
            commentsAndMarkup,
            bookmarks,
            optionalContent);
    }

    /// <summary>Low-level active-content and embedded-payload findings selected by the policy.</summary>
    public IReadOnlyList<PdfSanitizationFinding> Findings { get; }

    /// <summary>Per-kind action counts for the selected policy.</summary>
    public PdfSanitizationActionCounts ActionCounts { get; }

    /// <summary>Logical per-category counts for the selected sanitization policy.</summary>
    public PdfSanitizationCategoryCounts CategoryCounts { get; }

    /// <summary>Total number of low-level findings selected by the policy.</summary>
    public int TotalCount => Findings.Count;
}
