namespace OfficeIMO.Email.AddressBook;

/// <summary>Bounded, resumable query over decoded address-entry fields.</summary>
public sealed class OfflineAddressBookSearchQuery {
    /// <summary>Creates a bounded query.</summary>
    public OfflineAddressBookSearchQuery(
        IEnumerable<string> terms,
        OfflineAddressBookSearchFields fields = OfflineAddressBookSearchFields.All,
        OfflineAddressBookSearchMatchMode matchMode = OfflineAddressBookSearchMatchMode.AllTerms,
        string? addressListId = null,
        OfflineAddressBookObjectType? objectType = null,
        int maxEntriesScanned = 100_000,
        int maxResults = 100,
        int maxSearchableCharactersPerEntry = 1_000_000,
        int snippetCharacters = 240,
        int progressInterval = 1_000,
        bool continueOnEntryError = true,
        OfflineAddressBookSearchCheckpoint? resumeFrom = null) {
        if (terms == null) throw new ArgumentNullException(nameof(terms));
        const OfflineAddressBookSearchFields known = OfflineAddressBookSearchFields.All;
        if (fields == OfflineAddressBookSearchFields.None || (fields & ~known) != 0) {
            throw new ArgumentOutOfRangeException(nameof(fields));
        }
        if (!Enum.IsDefined(typeof(OfflineAddressBookSearchMatchMode), matchMode)) {
            throw new ArgumentOutOfRangeException(nameof(matchMode));
        }
        if (objectType.HasValue && !Enum.IsDefined(typeof(OfflineAddressBookObjectType), objectType.Value)) {
            throw new ArgumentOutOfRangeException(nameof(objectType));
        }
        if (maxEntriesScanned <= 0) throw new ArgumentOutOfRangeException(nameof(maxEntriesScanned));
        if (maxResults <= 0) throw new ArgumentOutOfRangeException(nameof(maxResults));
        if (maxSearchableCharactersPerEntry <= 0) {
            throw new ArgumentOutOfRangeException(nameof(maxSearchableCharactersPerEntry));
        }
        if (snippetCharacters < 32) throw new ArgumentOutOfRangeException(nameof(snippetCharacters));
        if (progressInterval <= 0) throw new ArgumentOutOfRangeException(nameof(progressInterval));

        string[] normalized = terms
            .Where(term => !string.IsNullOrWhiteSpace(term))
            .Select(term => term.Trim())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
        if (normalized.Length == 0 || normalized.Length > 32) {
            throw new ArgumentException("Address-book search requires between 1 and 32 non-empty terms.", nameof(terms));
        }
        if (normalized.Any(term => term.Length > 1024)) {
            throw new ArgumentException("An address-book search term cannot exceed 1,024 characters.", nameof(terms));
        }

        Terms = normalized;
        Fields = fields;
        MatchMode = matchMode;
        AddressListId = string.IsNullOrWhiteSpace(addressListId) ? null : addressListId;
        ObjectType = objectType;
        MaxEntriesScanned = maxEntriesScanned;
        MaxResults = maxResults;
        MaxSearchableCharactersPerEntry = maxSearchableCharactersPerEntry;
        SnippetCharacters = snippetCharacters;
        ProgressInterval = progressInterval;
        ContinueOnEntryError = continueOnEntryError;
        ResumeFrom = resumeFrom;
    }

    /// <summary>Case-insensitive terms.</summary>
    public IReadOnlyList<string> Terms { get; }
    /// <summary>Fields searched.</summary>
    public OfflineAddressBookSearchFields Fields { get; }
    /// <summary>Term-combination mode.</summary>
    public OfflineAddressBookSearchMatchMode MatchMode { get; }
    /// <summary>Optional address-list identifier.</summary>
    public string? AddressListId { get; }
    /// <summary>Optional projected object-type filter.</summary>
    public OfflineAddressBookObjectType? ObjectType { get; }
    /// <summary>Maximum records decoded in this batch.</summary>
    public int MaxEntriesScanned { get; }
    /// <summary>Maximum matches returned in this batch.</summary>
    public int MaxResults { get; }
    /// <summary>Maximum searchable characters retained across fields of one entry.</summary>
    public int MaxSearchableCharactersPerEntry { get; }
    /// <summary>Maximum snippet characters.</summary>
    public int SnippetCharacters { get; }
    /// <summary>Scanned-entry interval between progress reports.</summary>
    public int ProgressInterval { get; }
    /// <summary>Whether corrupt or over-limit records are diagnosed and skipped when their boundary is known.</summary>
    public bool ContinueOnEntryError { get; }
    /// <summary>Optional exact-position checkpoint from an earlier batch on the same session snapshot.</summary>
    public OfflineAddressBookSearchCheckpoint? ResumeFrom { get; }
}
