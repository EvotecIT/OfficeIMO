namespace OfficeIMO.Pdf;

/// <summary>Sanitized PDF bytes plus before/after proof and optional quarantined attachments.</summary>
public sealed class PdfSanitizationResult {
    private readonly byte[] _pdfBytes;
    private readonly PdfLoadOptions _readOptions;

    internal PdfSanitizationResult(
        byte[] pdfBytes,
        PdfMutationPlan mutationPlan,
        PdfRewritePreservationReport preservationReport,
        PdfSanitizationReport removedReport,
        PdfSanitizationReport remainingReport,
        IReadOnlyList<PdfExtractedAttachment> quarantinedAttachments,
        PdfLoadOptions readOptions) {
        _pdfBytes = (byte[])pdfBytes.Clone();
        _readOptions = readOptions;
        MutationPlan = mutationPlan;
        PreservationReport = preservationReport;
        RemovedReport = removedReport;
        RemainingReport = remainingReport;
        RemovedFindings = removedReport.Findings;
        RemainingFindings = remainingReport.Findings;
        QuarantinedAttachments = quarantinedAttachments;
        RemovedActionCounts = removedReport.ActionCounts;
        RemainingActionCounts = remainingReport.ActionCounts;
    }

    /// <summary>Shared mutation plan used for the full rewrite.</summary>
    public PdfMutationPlan MutationPlan { get; }

    /// <summary>Proof that document structures outside the sanitization policy were preserved.</summary>
    public PdfRewritePreservationReport PreservationReport { get; }

    /// <summary>Unsafe items present before the rewrite and removed by policy.</summary>
    public IReadOnlyList<PdfSanitizationFinding> RemovedFindings { get; }

    /// <summary>Forbidden items found after save. A successful operation always returns an empty list.</summary>
    public IReadOnlyList<PdfSanitizationFinding> RemainingFindings { get; }

    /// <summary>Typed inventory selected before the rewrite.</summary>
    public PdfSanitizationReport RemovedReport { get; }

    /// <summary>Typed selected inventory found after the rewrite.</summary>
    public PdfSanitizationReport RemainingReport { get; }

    /// <summary>Logical per-category counts removed by the policy.</summary>
    public PdfSanitizationCategoryCounts RemovedCategoryCounts => RemovedReport.CategoryCounts;

    /// <summary>Logical per-category counts still selected after the rewrite.</summary>
    public PdfSanitizationCategoryCounts RemainingCategoryCounts => RemainingReport.CategoryCounts;

    /// <summary>Per-kind counts of actions removed by the policy.</summary>
    public PdfSanitizationActionCounts RemovedActionCounts { get; }

    /// <summary>Per-kind counts of selected actions found after the rewrite.</summary>
    public PdfSanitizationActionCounts RemainingActionCounts { get; }

    /// <summary>Decoded attachments retained in memory when quarantine mode was requested.</summary>
    public IReadOnlyList<PdfExtractedAttachment> QuarantinedAttachments { get; }

    /// <summary>True when post-save inventory proves that no forbidden item remains.</summary>
    public bool IsSanitized => RemainingReport.CategoryCounts.Total == 0;

    /// <summary>Returns a defensive copy of the sanitized PDF bytes.</summary>
    public byte[] ToBytes() => (byte[])_pdfBytes.Clone();

    /// <summary>Opens the sanitized artifact as a fluent PDF document.</summary>
    public PdfDocument ToDocument() => PdfDocument.Load(_pdfBytes, _readOptions);
}
