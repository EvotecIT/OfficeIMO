namespace OfficeIMO.Email.Store;

/// <summary>Controls bounded lightweight item enumeration from an open store.</summary>
public sealed class EmailStoreEnumerationOptions {
    /// <summary>Creates item-enumeration options.</summary>
    public EmailStoreEnumerationOptions(
        string? folderId = null,
        bool includeDescendants = false,
        bool includeAssociatedItems = false,
        bool includeOrphanedItems = false,
        int maxItems = int.MaxValue) {
        if (maxItems <= 0) throw new ArgumentOutOfRangeException(nameof(maxItems));
        FolderId = string.IsNullOrWhiteSpace(folderId) ? null : folderId;
        IncludeDescendants = includeDescendants;
        IncludeAssociatedItems = includeAssociatedItems;
        IncludeOrphanedItems = includeOrphanedItems;
        MaxItems = maxItems;
    }

    /// <summary>Optional folder identifier. Null enumerates every folder.</summary>
    public string? FolderId { get; }

    /// <summary>Whether descendants of <see cref="FolderId"/> are included.</summary>
    public bool IncludeDescendants { get; }

    /// <summary>Whether folder-associated information items are included.</summary>
    public bool IncludeAssociatedItems { get; }

    /// <summary>Whether items missing from folder contents tables are recovered from source indexes.</summary>
    public bool IncludeOrphanedItems { get; }

    /// <summary>Maximum references returned by one enumeration.</summary>
    public int MaxItems { get; }
}
