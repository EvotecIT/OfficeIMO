namespace OfficeIMO.Email.Store;

/// <summary>An item reference and the lightweight summary that matched a store query.</summary>
public sealed class EmailStoreSearchResult {
    internal EmailStoreSearchResult(EmailStoreItemReference reference, EmailStoreItemSummary summary) {
        Reference = reference;
        Summary = summary;
    }

    /// <summary>Stable reference that can be passed to <see cref="EmailStoreSession.ReadItem(EmailStoreItemReference, CancellationToken)"/>.</summary>
    public EmailStoreItemReference Reference { get; }

    /// <summary>Lightweight summary that satisfied the query.</summary>
    public EmailStoreItemSummary Summary { get; }
}
