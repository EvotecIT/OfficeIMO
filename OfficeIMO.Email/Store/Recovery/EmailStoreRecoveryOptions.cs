namespace OfficeIMO.Email.Store;

/// <summary>Bounds for non-mutating recovery discovery through store indexes.</summary>
public sealed class EmailStoreRecoveryOptions {
    /// <summary>Creates recovery-discovery options.</summary>
    public EmailStoreRecoveryOptions(
        string? folderId = null,
        bool includeDescendants = false,
        bool includeAssociatedItems = false,
        int maxItemsScanned = 1_000_000,
        int maxRecoveredItems = 10_000) {
        if (maxItemsScanned <= 0) throw new ArgumentOutOfRangeException(nameof(maxItemsScanned));
        if (maxRecoveredItems <= 0) throw new ArgumentOutOfRangeException(nameof(maxRecoveredItems));
        FolderId = string.IsNullOrWhiteSpace(folderId) ? null : folderId;
        IncludeDescendants = includeDescendants;
        IncludeAssociatedItems = includeAssociatedItems;
        MaxItemsScanned = maxItemsScanned;
        MaxRecoveredItems = maxRecoveredItems;
    }

    /// <summary>Optional folder identifier. Null scans every folder.</summary>
    public string? FolderId { get; }

    /// <summary>Whether descendants of <see cref="FolderId"/> are included.</summary>
    public bool IncludeDescendants { get; }

    /// <summary>Whether associated-item indexes are scanned.</summary>
    public bool IncludeAssociatedItems { get; }

    /// <summary>Maximum item references examined.</summary>
    public int MaxItemsScanned { get; }

    /// <summary>Maximum recovered references retained in the report.</summary>
    public int MaxRecoveredItems { get; }
}
