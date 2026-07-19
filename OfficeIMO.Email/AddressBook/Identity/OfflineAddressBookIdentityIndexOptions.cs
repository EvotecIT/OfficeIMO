namespace OfficeIMO.Email.AddressBook;

/// <summary>Bounds and matching policy used while building an offline directory identity index.</summary>
public sealed class OfflineAddressBookIdentityIndexOptions {
    /// <summary>Creates identity-index options.</summary>
    public OfflineAddressBookIdentityIndexOptions(
        string? addressListId = null,
        int maxEntries = 250000,
        int maxIdentitiesPerEntry = 128,
        bool includeAccountNames = true,
        bool includeDisplayNames = false,
        bool continueOnEntryError = true) {
        if (maxEntries <= 0) throw new ArgumentOutOfRangeException(nameof(maxEntries));
        if (maxIdentitiesPerEntry <= 0) throw new ArgumentOutOfRangeException(nameof(maxIdentitiesPerEntry));
        AddressListId = string.IsNullOrWhiteSpace(addressListId) ? null : addressListId;
        MaxEntries = maxEntries;
        MaxIdentitiesPerEntry = maxIdentitiesPerEntry;
        IncludeAccountNames = includeAccountNames;
        IncludeDisplayNames = includeDisplayNames;
        ContinueOnEntryError = continueOnEntryError;
    }

    /// <summary>Optional address-list identifier. Null indexes every discovered Full Details list.</summary>
    public string? AddressListId { get; }

    /// <summary>Maximum successfully decoded entries included in the index.</summary>
    public int MaxEntries { get; }

    /// <summary>Maximum distinct lookup identities retained for one directory entry.</summary>
    public int MaxIdentitiesPerEntry { get; }

    /// <summary>Whether exact directory account names are indexed as non-address aliases.</summary>
    public bool IncludeAccountNames { get; }

    /// <summary>
    /// Whether display names are indexed. They remain heuristic and must also be enabled by each query.
    /// </summary>
    public bool IncludeDisplayNames { get; }

    /// <summary>Whether safely framed corrupt records are diagnosed and skipped.</summary>
    public bool ContinueOnEntryError { get; }
}
