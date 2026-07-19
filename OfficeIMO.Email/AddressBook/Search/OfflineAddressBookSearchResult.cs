namespace OfficeIMO.Email.AddressBook;

/// <summary>One bounded search match.</summary>
public sealed class OfflineAddressBookSearchResult {
    internal OfflineAddressBookSearchResult(OfflineAddressBookEntrySummary summary,
        OfflineAddressBookSearchFields matchedFields, string? snippet) {
        Summary = summary;
        MatchedFields = matchedFields;
        Snippet = snippet;
    }

    /// <summary>Small typed projection with a stable reference for an explicit full read.</summary>
    public OfflineAddressBookEntrySummary Summary { get; }
    /// <summary>Selected fields containing at least one term.</summary>
    public OfflineAddressBookSearchFields MatchedFields { get; }
    /// <summary>Bounded text around the earliest match.</summary>
    public string? Snippet { get; }
}
