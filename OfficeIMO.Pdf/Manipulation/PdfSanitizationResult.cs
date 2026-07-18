namespace OfficeIMO.Pdf;

/// <summary>Sanitized PDF bytes plus before/after proof and optional quarantined attachments.</summary>
public sealed class PdfSanitizationResult {
    private readonly byte[] _pdfBytes;

    internal PdfSanitizationResult(
        byte[] pdfBytes,
        PdfMutationPlan mutationPlan,
        IReadOnlyList<PdfSanitizationFinding> removedFindings,
        IReadOnlyList<PdfSanitizationFinding> remainingFindings,
        IReadOnlyList<PdfExtractedAttachment> quarantinedAttachments) {
        _pdfBytes = (byte[])pdfBytes.Clone();
        MutationPlan = mutationPlan;
        RemovedFindings = removedFindings;
        RemainingFindings = remainingFindings;
        QuarantinedAttachments = quarantinedAttachments;
    }

    /// <summary>Shared mutation plan used for the full rewrite.</summary>
    public PdfMutationPlan MutationPlan { get; }

    /// <summary>Unsafe items present before the rewrite and removed by policy.</summary>
    public IReadOnlyList<PdfSanitizationFinding> RemovedFindings { get; }

    /// <summary>Forbidden items found after save. A successful operation always returns an empty list.</summary>
    public IReadOnlyList<PdfSanitizationFinding> RemainingFindings { get; }

    /// <summary>Decoded attachments retained in memory when quarantine mode was requested.</summary>
    public IReadOnlyList<PdfExtractedAttachment> QuarantinedAttachments { get; }

    /// <summary>True when post-save inventory proves that no forbidden item remains.</summary>
    public bool IsSanitized => RemainingFindings.Count == 0;

    /// <summary>Returns a defensive copy of the sanitized PDF bytes.</summary>
    public byte[] ToBytes() => (byte[])_pdfBytes.Clone();

    /// <summary>Opens the sanitized artifact as a fluent PDF document.</summary>
    public PdfDocument ToDocument() => PdfDocument.Open(_pdfBytes);
}
