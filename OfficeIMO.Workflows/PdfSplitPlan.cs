namespace OfficeIMO.Workflows;

/// <summary>One planned consecutive output part.</summary>
public sealed record PdfSplitPart(string Name, int FirstSourcePage, int PageCount);

/// <summary>A page where a new output part starts, and the title used to name that part.</summary>
public sealed record PdfSplitStart(int FirstPage, string Title);

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

    /// <summary>
    /// Creates a plan that starts a new part at each supplied page and names parts after their titles.
    /// Pages before the first start become an unnamed leading part; duplicate starts are merged.
    /// </summary>
    public static PdfSplitPlan FromStarts(int pageCount, IReadOnlyList<PdfSplitStart> starts, int maximumParts = 1000) {
        if (pageCount < 1) throw new ArgumentOutOfRangeException(nameof(pageCount));
        if (starts is null) throw new ArgumentNullException(nameof(starts));
        if (maximumParts < 1) throw new ArgumentOutOfRangeException(nameof(maximumParts));
        var ordered = starts.Where(start => start is not null && start.FirstPage >= 1 && start.FirstPage <= pageCount)
            .GroupBy(start => start.FirstPage).Select(group => group.First()).OrderBy(start => start.FirstPage).ToList();
        if (ordered.Count == 0) throw new ArgumentException("No split starts fall inside the document.", nameof(starts));
        if (ordered[0].FirstPage > 1) ordered.Insert(0, new PdfSplitStart(1, string.Empty));
        if (ordered.Count > maximumParts) throw new InvalidOperationException($"The split would create {ordered.Count} parts, above the configured {maximumParts}-part limit.");
        var parts = new PdfSplitPart[ordered.Count];
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        for (int index = 0; index < ordered.Count; index++) {
            int first = ordered[index].FirstPage;
            int last = index + 1 < ordered.Count ? ordered[index + 1].FirstPage - 1 : pageCount;
            string title = SafeTitle(ordered[index].Title);
            string name = title.Length == 0 ? $"part-{index + 1:D3}.pdf" : $"{index + 1:D3}-{title}.pdf";
            if (!names.Add(name)) name = $"part-{index + 1:D3}.pdf";
            parts[index] = new(name, first, last - first + 1);
        }
        return new(parts);
    }

    /// <summary>Checks that named parts cover every source page exactly once in order.</summary>
    internal void Validate(int pageCount, int maximumParts) {
        if (Parts.Count == 0) throw new ArgumentException("The split plan has no parts.");
        if (Parts.Count > maximumParts) throw new InvalidOperationException($"The split would create {Parts.Count} parts, above the configured {maximumParts}-part limit.");
        long next = 1;
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (PdfSplitPart part in Parts) {
            if (part.FirstSourcePage != next || part.PageCount < 1 ||
                (long)part.FirstSourcePage + part.PageCount - 1 > pageCount)
                throw new ArgumentException("Split parts must cover every source page exactly once in order.");
            if (string.IsNullOrWhiteSpace(part.Name) || part.Name != Path.GetFileName(part.Name) || !part.Name.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase) ||
                part.Name.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 || !names.Add(part.Name))
                throw new ArgumentException("Split part names must be unique PDF file names.");
            next = (long)part.FirstSourcePage + part.PageCount;
        }
        if (next != pageCount + 1L)
            throw new ArgumentException("Split parts must cover every source page exactly once in order.");
    }

    private static string SafeTitle(string? title) {
        if (string.IsNullOrWhiteSpace(title)) return string.Empty;
        var invalid = new HashSet<char>(Path.GetInvalidFileNameChars().Concat(['/', '\\', ':', '*', '?', '"', '<', '>', '|']));
        string cleaned = new string(title.Trim().Select(character => invalid.Contains(character) || char.IsControl(character) ? ' ' : character).ToArray());
        cleaned = string.Join(" ", cleaned.Split(' ', StringSplitOptions.RemoveEmptyEntries)).Trim('.', ' ');
        return cleaned.Length > 60 ? cleaned[..60].TrimEnd('.', ' ') : cleaned;
    }
}
