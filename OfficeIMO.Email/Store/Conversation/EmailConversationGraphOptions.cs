namespace OfficeIMO.Email.Store;

/// <summary>Scope, bounds, and heuristic policy for offline conversation graph construction.</summary>
public sealed class EmailConversationGraphOptions {
    /// <summary>Creates conversation graph options.</summary>
    public EmailConversationGraphOptions(
        EmailStoreFolderId? folderId = null,
        bool includeDescendants = true,
        bool includeAssociatedItems = false,
        bool includeOrphanedItems = true,
        bool includeSubjectHeuristics = true,
        bool includeMeetingAndTaskLinks = true,
        bool continueOnItemError = true,
        int maxItems = 100000,
        int maxEdges = 500000,
        int maxReferencesPerItem = 256,
        long maxDecodedPropertyBytesPerItem = 4L * 1024L * 1024L) {
        if (folderId.HasValue && folderId.Value.IsEmpty) throw new ArgumentException(
            "A folder scope cannot be empty.", nameof(folderId));
        if (maxItems <= 0) throw new ArgumentOutOfRangeException(nameof(maxItems));
        if (maxEdges <= 0) throw new ArgumentOutOfRangeException(nameof(maxEdges));
        if (maxReferencesPerItem <= 0) throw new ArgumentOutOfRangeException(nameof(maxReferencesPerItem));
        if (maxDecodedPropertyBytesPerItem <= 0) {
            throw new ArgumentOutOfRangeException(nameof(maxDecodedPropertyBytesPerItem));
        }
        FolderId = folderId;
        IncludeDescendants = includeDescendants;
        IncludeAssociatedItems = includeAssociatedItems;
        IncludeOrphanedItems = includeOrphanedItems;
        IncludeSubjectHeuristics = includeSubjectHeuristics;
        IncludeMeetingAndTaskLinks = includeMeetingAndTaskLinks;
        ContinueOnItemError = continueOnItemError;
        MaxItems = maxItems;
        MaxEdges = maxEdges;
        MaxReferencesPerItem = maxReferencesPerItem;
        MaxDecodedPropertyBytesPerItem = maxDecodedPropertyBytesPerItem;
    }

    /// <summary>Optional typed folder scope. Null scans every folder.</summary>
    public EmailStoreFolderId? FolderId { get; }
    /// <summary>Whether descendants of the selected folder are included.</summary>
    public bool IncludeDescendants { get; }
    /// <summary>Whether folder-associated information items are included.</summary>
    public bool IncludeAssociatedItems { get; }
    /// <summary>Whether indexed items missing from folder contents tables are included.</summary>
    public bool IncludeOrphanedItems { get; }
    /// <summary>
    /// Whether otherwise unconnected items may be related by exact conversation topic or normalized subject.
    /// Such links are always marked heuristic.
    /// </summary>
    public bool IncludeSubjectHeuristics { get; }
    /// <summary>Whether meeting Global Object IDs and task Global IDs form strong related-item links.</summary>
    public bool IncludeMeetingAndTaskLinks { get; }
    /// <summary>Whether an unreadable item is diagnosed and retained as a summary-only graph node.</summary>
    public bool ContinueOnItemError { get; }
    /// <summary>Maximum items scanned. The builder probes one additional reference to report truncation.</summary>
    public int MaxItems { get; }
    /// <summary>Maximum distinct graph edges.</summary>
    public int MaxEdges { get; }
    /// <summary>Maximum Message-ID values decoded from one References field.</summary>
    public int MaxReferencesPerItem { get; }
    /// <summary>Per-item decoded-property byte cap for selective metadata reads.</summary>
    public long MaxDecodedPropertyBytesPerItem { get; }
}
