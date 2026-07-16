namespace OfficeIMO.Email.Store;

/// <summary>Bounded, resumable query over summary and selected item content.</summary>
public sealed class EmailStoreContentQuery {
    /// <summary>Creates a bounded content query.</summary>
    public EmailStoreContentQuery(
        IEnumerable<string> terms,
        EmailStoreContentSearchFields fields = EmailStoreContentSearchFields.All,
        EmailStoreContentMatchMode matchMode = EmailStoreContentMatchMode.AllTerms,
        EmailStoreQuery? metadataFilter = null,
        int maxItemsScanned = 10_000,
        int maxResults = 100,
        long maxDecodedPropertyBytesPerItem = 16L * 1024 * 1024,
        int maxSearchableCharactersPerItem = 2_000_000,
        int snippetCharacters = 240,
        int progressInterval = 100,
        bool continueOnItemError = true,
        EmailStoreContentSearchCheckpoint? resumeFrom = null) {
        if (terms == null) throw new ArgumentNullException(nameof(terms));
        const EmailStoreContentSearchFields known = EmailStoreContentSearchFields.All;
        if (fields == EmailStoreContentSearchFields.None || (fields & ~known) != 0) {
            throw new ArgumentOutOfRangeException(nameof(fields));
        }
        if (!Enum.IsDefined(typeof(EmailStoreContentMatchMode), matchMode)) {
            throw new ArgumentOutOfRangeException(nameof(matchMode));
        }
        if (maxItemsScanned <= 0) throw new ArgumentOutOfRangeException(nameof(maxItemsScanned));
        if (maxResults <= 0) throw new ArgumentOutOfRangeException(nameof(maxResults));
        if (maxDecodedPropertyBytesPerItem <= 0) {
            throw new ArgumentOutOfRangeException(nameof(maxDecodedPropertyBytesPerItem));
        }
        if (maxSearchableCharactersPerItem <= 0) {
            throw new ArgumentOutOfRangeException(nameof(maxSearchableCharactersPerItem));
        }
        if (snippetCharacters < 32) throw new ArgumentOutOfRangeException(nameof(snippetCharacters));
        if (progressInterval <= 0) throw new ArgumentOutOfRangeException(nameof(progressInterval));

        string[] normalized = terms
            .Where(term => !string.IsNullOrWhiteSpace(term))
            .Select(term => term.Trim())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
        if (normalized.Length == 0 || normalized.Length > 32) {
            throw new ArgumentException("Content search requires between 1 and 32 non-empty terms.", nameof(terms));
        }
        if (normalized.Any(term => term.Length > 1024)) {
            throw new ArgumentException("A content-search term cannot exceed 1,024 characters.", nameof(terms));
        }

        Terms = normalized;
        Fields = fields;
        MatchMode = matchMode;
        MetadataFilter = metadataFilter;
        MaxItemsScanned = maxItemsScanned;
        MaxResults = maxResults;
        MaxDecodedPropertyBytesPerItem = maxDecodedPropertyBytesPerItem;
        MaxSearchableCharactersPerItem = maxSearchableCharactersPerItem;
        SnippetCharacters = snippetCharacters;
        ProgressInterval = progressInterval;
        ContinueOnItemError = continueOnItemError;
        ResumeFrom = resumeFrom;
    }

    /// <summary>Case-insensitive terms.</summary>
    public IReadOnlyList<string> Terms { get; }
    /// <summary>Fields searched.</summary>
    public EmailStoreContentSearchFields Fields { get; }
    /// <summary>Term-combination mode.</summary>
    public EmailStoreContentMatchMode MatchMode { get; }
    /// <summary>Optional cheap summary filter evaluated before item content is decoded.</summary>
    public EmailStoreQuery? MetadataFilter { get; }
    /// <summary>Maximum item references processed in this batch.</summary>
    public int MaxItemsScanned { get; }
    /// <summary>Maximum matches returned in this batch.</summary>
    public int MaxResults { get; }
    /// <summary>Maximum decoded property bytes allowed for one selective item read.</summary>
    public long MaxDecodedPropertyBytesPerItem { get; }
    /// <summary>Maximum searchable characters retained across fields of one item.</summary>
    public int MaxSearchableCharactersPerItem { get; }
    /// <summary>Maximum snippet characters.</summary>
    public int SnippetCharacters { get; }
    /// <summary>Scanned-item interval between progress notifications.</summary>
    public int ProgressInterval { get; }
    /// <summary>Whether corrupt or over-limit items are diagnosed and skipped.</summary>
    public bool ContinueOnItemError { get; }
    /// <summary>Optional checkpoint from a previous batch.</summary>
    public EmailStoreContentSearchCheckpoint? ResumeFrom { get; }
}
