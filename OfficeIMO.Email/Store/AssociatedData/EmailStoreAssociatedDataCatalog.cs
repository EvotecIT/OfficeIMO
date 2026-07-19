namespace OfficeIMO.Email.Store;

/// <summary>Bounded typed catalog of Store folder-associated information.</summary>
public sealed class EmailStoreAssociatedDataCatalog {
    internal EmailStoreAssociatedDataCatalog(IReadOnlyList<EmailStoreAssociatedItem> items,
        IReadOnlyList<EmailStoreFolderInfo> folders,
        IReadOnlyList<EmailStoreDiagnostic> diagnostics, bool isComplete, int scannedItems) {
        Items = items;
        Diagnostics = diagnostics;
        IsComplete = isComplete;
        ScannedItems = scannedItems;
        CategoryLists = Filter(item => item.CategoryList != null);
        Configurations = Filter(item => item.Configuration != null);
        Views = Filter(item => item.ViewDefinition != null);
        RuleOrganizers = Filter(item => item.RuleOrganizer != null);
        SearchFolders = Filter(item => item.SearchFolderDefinition != null);
        FolderUserProperties = Filter(item => item.FolderUserProperties != null);
        SearchFolderContainers = folders.Select(folder => new EmailStoreSearchFolderContainer(folder))
            .Where(container => container.HasSearchEvidence).ToArray();
    }

    /// <summary>Every successfully read associated message in source order.</summary>
    public IReadOnlyList<EmailStoreAssociatedItem> Items { get; }

    /// <summary>Messages carrying parsed category lists.</summary>
    public IReadOnlyList<EmailStoreAssociatedItem> CategoryLists { get; }

    /// <summary>Messages carrying roaming configuration data.</summary>
    public IReadOnlyList<EmailStoreAssociatedItem> Configurations { get; }

    /// <summary>Named-view definition messages.</summary>
    public IReadOnlyList<EmailStoreAssociatedItem> Views { get; }

    /// <summary>Rule organizer messages.</summary>
    public IReadOnlyList<EmailStoreAssociatedItem> RuleOrganizers { get; }

    /// <summary>Persistent search-folder definition messages.</summary>
    public IReadOnlyList<EmailStoreAssociatedItem> SearchFolders { get; }

    /// <summary>Associated messages carrying Outlook field definitions.</summary>
    public IReadOnlyList<EmailStoreAssociatedItem> FolderUserProperties { get; }

    /// <summary>Search-folder containers projected from folder-owned MAPI properties.</summary>
    public IReadOnlyList<EmailStoreSearchFolderContainer> SearchFolderContainers { get; }

    /// <summary>Catalog- and item-level diagnostics.</summary>
    public IReadOnlyList<EmailStoreDiagnostic> Diagnostics { get; }

    /// <summary>True when every item in scope was read within bounds.</summary>
    public bool IsComplete { get; }

    /// <summary>Number of associated references examined.</summary>
    public int ScannedItems { get; }

    /// <summary>
    /// Returns the most recently modified configuration message for one class and folder. Duplicate candidates
    /// remain visible through <see cref="Configurations"/> and are reported as a conflict diagnostic.
    /// </summary>
    public EmailStoreAssociatedItem? FindEffectiveConfiguration(string messageClass,
        EmailStoreFolderId? folderId = null) {
        if (string.IsNullOrWhiteSpace(messageClass)) throw new ArgumentException("Message class is required.", nameof(messageClass));
        return Configurations
            .Where(item => string.Equals(item.Document.MessageClass, messageClass,
                StringComparison.OrdinalIgnoreCase) &&
                (!folderId.HasValue || item.Folder.Key == folderId.Value))
            .OrderByDescending(item => item.ModifiedAt ?? DateTimeOffset.MinValue)
            .ThenByDescending(item => item.Reference.Id, StringComparer.Ordinal)
            .FirstOrDefault();
    }

    private IReadOnlyList<EmailStoreAssociatedItem> Filter(
        Func<EmailStoreAssociatedItem, bool> predicate) =>
        Items.Where(predicate).ToArray();
}
