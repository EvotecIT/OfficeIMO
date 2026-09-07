namespace OfficeIMO.Workflows;

/// <summary>One planned consecutive output part.</summary>
public sealed record PdfSplitPart(string Name, int FirstSourcePage, int PageCount);

/// <summary>Deterministic split filenames and page ranges shared by previews and execution.</summary>
public sealed class PdfSplitPlan {
    private PdfSplitPlan(PdfSplitPart[] parts) => Parts = Array.AsReadOnly(parts);
    /// <summary>Output parts in source page order.</summary>
    public IReadOnlyList<PdfSplitPart> Parts { get; }
    /// <summary>Creates a bounded plan without loading or changing a document.</summary>
    public static PdfSplitPlan Create(int pageCount, int pagesPerDocument, int maximumParts = 1000) {
        if (pageCount < 1) throw new ArgumentOutOfRangeException(nameof(pageCount));
        if (pagesPerDocument < 1) throw new ArgumentOutOfRangeException(nameof(pagesPerDocument));
        if (maximumParts < 1) throw new ArgumentOutOfRangeException(nameof(maximumParts));
        int count = 1 + (pageCount - 1) / pagesPerDocument;
        if (count > maximumParts) throw new InvalidOperationException($"The split would create {count} parts, above the configured {maximumParts}-part limit.");
        var parts = new PdfSplitPart[count];
        for (int index = 0; index < count; index++) {
            int first = index * pagesPerDocument + 1;
            parts[index] = new($"part-{index + 1:D3}.pdf", first, Math.Min(pagesPerDocument, pageCount - first + 1));
        }
        return new(parts);
    }
}
