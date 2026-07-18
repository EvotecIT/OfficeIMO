namespace OfficeIMO.Email.Store;

/// <summary>Controls bounded lightweight item enumeration from an open store.</summary>
public sealed class EmailStoreEnumerationOptions {
    /// <summary>Creates item-enumeration options.</summary>
    public EmailStoreEnumerationOptions(
        string? folderId = null,
        bool includeDescendants = false,
        bool includeAssociatedItems = false,
        bool includeOrphanedItems = false,
        int maxItems = int.MaxValue,
        bool includeRegularItems = true) {
        if (maxItems <= 0) throw new ArgumentOutOfRangeException(nameof(maxItems));
        if (!includeRegularItems && !includeAssociatedItems) {
            throw new ArgumentException("At least one regular or associated item scope must be included.");
        }
        FolderId = string.IsNullOrWhiteSpace(folderId) ? null : folderId;
        IncludeDescendants = includeDescendants;
        IncludeAssociatedItems = includeAssociatedItems;
        IncludeOrphanedItems = includeOrphanedItems;
        MaxItems = maxItems;
        IncludeRegularItems = includeRegularItems;
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

    /// <summary>Whether normal folder contents are included.</summary>
    public bool IncludeRegularItems { get; }

    /// <summary>Creates options for one typed folder scope.</summary>
    public static EmailStoreEnumerationOptions ForFolder(EmailStoreFolderId folderId,
        bool includeDescendants = false, bool includeAssociatedItems = false,
        int maxItems = int.MaxValue) =>
        new EmailStoreEnumerationOptions(folderId.Value, includeDescendants,
            includeAssociatedItems, includeOrphanedItems: false, maxItems);

    /// <summary>Creates an associated-information-only scope without scanning normal item tables.</summary>
    public static EmailStoreEnumerationOptions ForAssociated(
        EmailStoreFolderId? folderId = null, bool includeDescendants = false,
        int maxItems = int.MaxValue) =>
        new EmailStoreEnumerationOptions(folderId?.Value, includeDescendants,
            includeAssociatedItems: true, includeOrphanedItems: false,
            maxItems: maxItems, includeRegularItems: false);
}
