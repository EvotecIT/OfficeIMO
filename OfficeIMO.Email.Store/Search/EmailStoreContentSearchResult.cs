namespace OfficeIMO.Email.Store;

/// <summary>One content-search match with its summary, matched fields, and bounded snippet.</summary>
public sealed class EmailStoreContentSearchResult {
    internal EmailStoreContentSearchResult(EmailStoreItemReference reference,
        EmailStoreItemSummary summary, EmailStoreContentSearchFields matchedFields,
        string? snippet) {
        Reference = reference;
        Summary = summary;
        MatchedFields = matchedFields;
        Snippet = snippet;
    }

    /// <summary>Stable reference for an explicit selective or full item read.</summary>
    public EmailStoreItemReference Reference { get; }
    /// <summary>Lightweight item summary.</summary>
    public EmailStoreItemSummary Summary { get; }
    /// <summary>Selected fields containing at least one query term.</summary>
    public EmailStoreContentSearchFields MatchedFields { get; }
    /// <summary>Bounded text around the earliest match, when text was available.</summary>
    public string? Snippet { get; }
}
