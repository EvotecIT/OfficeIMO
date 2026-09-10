namespace OfficeIMO.AI;

/// <summary>Local text-evidence availability. It does not certify recognition quality or model reasoning.</summary>
public sealed record OfficeAiEvidenceReadiness {
    /// <summary>Known pages in the requested scope, in document order.</summary>
    public required IReadOnlyList<int> Pages { get; init; }
    /// <summary>Known scoped pages without any extractable text observations.</summary>
    public required IReadOnlyList<int> PagesWithoutText { get; init; }
    /// <summary>Number of nonempty text observations in scope.</summary>
    public int TextItems { get; init; }
    /// <summary>Total characters in the scoped observations, before model request limits.</summary>
    public long TextCharacters { get; init; }
    /// <summary>Whether any text can be supplied to a text-only operation.</summary>
    public bool HasText => TextItems > 0;
    /// <summary>Whether the reader reported limitations when preparing this snapshot.</summary>
    public bool HasSourceDiagnostics { get; init; }

    /// <summary>Inspects already prepared immutable evidence without contacting a provider or re-reading the source.</summary>
    public static OfficeAiEvidenceReadiness Inspect(OfficeAiDocument document, IEnumerable<int>? pages = null) {
        ArgumentNullException.ThrowIfNull(document);
        int[] selected = pages?.Distinct().OrderBy(page => page).ToArray() ?? [];
        if (selected.Any(page => page < 1 || !document.Pages.Contains(page)))
            throw new ArgumentOutOfRangeException(nameof(pages), "The selected page is not present in this snapshot.");
        bool scoped = selected.Length > 0;
        int[] knownPages = scoped ? selected : document.Pages.ToArray();
        var scope = selected.ToHashSet();
        OfficeAiEvidence[] text = document.Evidence.Where(item => !string.IsNullOrWhiteSpace(item.Text)
            && (!scoped || item.Page.HasValue && scope.Contains(item.Page.Value))).ToArray();
        var textPages = text.Where(item => item.Page.HasValue).Select(item => item.Page!.Value).ToHashSet();
        return new() {
            Pages = Array.AsReadOnly(knownPages),
            PagesWithoutText = Array.AsReadOnly(knownPages.Where(page => !textPages.Contains(page)).ToArray()),
            TextItems = text.Length, TextCharacters = text.Sum(item => (long)item.Text.Length),
            HasSourceDiagnostics = document.HasSourceDiagnostics
        };
    }
}
