namespace OfficeIMO.Email.Store;

/// <summary>A folder in an email store.</summary>
public sealed class EmailStoreFolder {
    private readonly List<EmailStoreItem> _items = new List<EmailStoreItem>();
    private readonly List<EmailStoreItem> _associatedItems = new List<EmailStoreItem>();

    internal EmailStoreFolder(string id, string? parentId, string name) {
        Id = id;
        ParentId = parentId;
        Name = name;
    }

    /// <summary>Stable source identifier.</summary>
    public string Id { get; }

    /// <summary>Parent folder identifier, or null for a root.</summary>
    public string? ParentId { get; }

    /// <summary>Display name.</summary>
    public string Name { get; internal set; }

    /// <summary>Items directly contained in this folder.</summary>
    public IReadOnlyList<EmailStoreItem> Items => _items;

    /// <summary>Folder-associated information items when explicitly requested by reader options.</summary>
    public IReadOnlyList<EmailStoreItem> AssociatedItems => _associatedItems;

    internal IList<EmailStoreItem> MutableItems => _items;
    internal IList<EmailStoreItem> MutableAssociatedItems => _associatedItems;
}
