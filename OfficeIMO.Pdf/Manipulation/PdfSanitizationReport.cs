namespace OfficeIMO.Pdf;

/// <summary>Typed preview of items that the supplied PDF sanitization policy would remove.</summary>
internal sealed class PdfSanitizationReport {
    internal PdfSanitizationReport(IReadOnlyList<PdfSanitizationFinding> findings, IReadOnlyList<PdfSanitizationItem>? items = null) {
        Findings = findings;
        Items = items ?? Array.Empty<PdfSanitizationItem>();
        ActionCounts = new PdfSanitizationActionCounts(findings);
        CategoryCounts = new PdfSanitizationCategoryCounts(
            Count(PdfSanitizationContentKind.UserMetadata),
            Count(PdfSanitizationContentKind.EmbeddedFiles),
            ActionCounts.Total,
            Count(PdfSanitizationContentKind.CommentsAndMarkup),
            Count(PdfSanitizationContentKind.Bookmarks),
            Count(PdfSanitizationContentKind.OptionalContent));
    }

    /// <summary>Low-level active-content and embedded-payload findings selected by the policy.</summary>
    public IReadOnlyList<PdfSanitizationFinding> Findings { get; }

    /// <summary>Logical metadata, attachment, comment, bookmark and layer items selected by the policy, including the ones removal keeps on purpose.</summary>
    public IReadOnlyList<PdfSanitizationItem> Items { get; }

    /// <summary>Per-kind action counts for the selected policy.</summary>
    public PdfSanitizationActionCounts ActionCounts { get; }

    /// <summary>Logical per-category counts for the selected sanitization policy.</summary>
    public PdfSanitizationCategoryCounts CategoryCounts { get; }

    /// <summary>Total number of low-level findings selected by the policy.</summary>
    public int TotalCount => Findings.Count;

    private int Count(PdfSanitizationContentKind kind) {
        int count = 0;
        for (int index = 0; index < Items.Count; index++) {
            if (Items[index].Kind == kind && !Items[index].IsRetained) count++;
        }
        return count;
    }
}

/// <summary>One logical item in a non-action sanitization category.</summary>
internal sealed class PdfSanitizationItem {
    internal PdfSanitizationItem(PdfSanitizationContentKind kind, string location, string description, string? value, bool isProvenanceCarrier = false,
        string? retainedNote = null) {
        Kind = kind;
        Location = location;
        Description = description;
        Value = value;
        IsProvenanceCarrier = isProvenanceCarrier;
        RetainedNote = retainedNote;
    }

    internal PdfSanitizationContentKind Kind { get; }
    internal string Location { get; }
    internal string Description { get; }
    internal string? Value { get; }

    /// <summary>An attachment that carries Content Credentials, reported as provenance rather than as an attachment.</summary>
    internal bool IsProvenanceCarrier { get; }

    /// <summary>Why removal keeps this item even when its category is selected, such as the document title; null for items removal takes out.</summary>
    internal string? RetainedNote { get; }

    internal bool IsRetained => RetainedNote is not null;
}
