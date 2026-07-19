namespace OfficeIMO.Email.Store;

/// <summary>Bounds and domain behavior for building an Outlook-compatible reminder queue.</summary>
public sealed class EmailStoreReminderQueryOptions {
    /// <summary>Creates immutable reminder query options.</summary>
    public EmailStoreReminderQueryOptions(
        EmailStoreFolderId? folderId = null,
        bool includeDescendants = false,
        bool includeInactive = false,
        bool includeExcludedFolders = false,
        DateTimeOffset? asOf = null,
        int maxItemsScanned = 1_000_000,
        int maxResults = 100_000,
        long maxDecodedPropertyBytesPerItem = 4 * 1024 * 1024,
        bool continueOnError = true) {
        if (folderId.HasValue && folderId.Value.IsEmpty) throw new ArgumentException("The folder identifier cannot be empty.", nameof(folderId));
        if (maxItemsScanned <= 0) throw new ArgumentOutOfRangeException(nameof(maxItemsScanned));
        if (maxResults <= 0) throw new ArgumentOutOfRangeException(nameof(maxResults));
        if (maxDecodedPropertyBytesPerItem <= 0) throw new ArgumentOutOfRangeException(nameof(maxDecodedPropertyBytesPerItem));
        FolderId = folderId;
        IncludeDescendants = includeDescendants;
        IncludeInactive = includeInactive;
        IncludeExcludedFolders = includeExcludedFolders;
        AsOf = (asOf ?? DateTimeOffset.UtcNow).ToUniversalTime();
        MaxItemsScanned = maxItemsScanned;
        MaxResults = maxResults;
        MaxDecodedPropertyBytesPerItem = maxDecodedPropertyBytesPerItem;
        ContinueOnError = continueOnError;
    }

    /// <summary>Optional folder scope.</summary>
    public EmailStoreFolderId? FolderId { get; }
    /// <summary>Whether descendants of the selected folder are included.</summary>
    public bool IncludeDescendants { get; }
    /// <summary>Whether disabled reminders with reminder-property evidence are returned.</summary>
    public bool IncludeInactive { get; }
    /// <summary>Whether folders excluded from the Outlook reminder domain are scanned.</summary>
    public bool IncludeExcludedFolders { get; }
    /// <summary>UTC instant used to classify pending and overdue reminders.</summary>
    public DateTimeOffset AsOf { get; }
    /// <summary>Maximum eligible normal item references examined after reminder-domain folder filtering.</summary>
    public int MaxItemsScanned { get; }
    /// <summary>Maximum reminder rows returned.</summary>
    public int MaxResults { get; }
    /// <summary>Maximum decoded root-property bytes per examined item.</summary>
    public long MaxDecodedPropertyBytesPerItem { get; }
    /// <summary>Whether corrupt individual items are diagnosed and skipped.</summary>
    public bool ContinueOnError { get; }
}
