namespace OfficeIMO.Email.Store;

/// <summary>Lightweight folder metadata exposed by an open email-store session.</summary>
public sealed class EmailStoreFolderInfo {
    internal EmailStoreFolderInfo(string id, string? parentId, string name,
        int? itemCount = null, int? associatedItemCount = null) {
        Id = id ?? throw new ArgumentNullException(nameof(id));
        ParentId = parentId;
        Name = name ?? throw new ArgumentNullException(nameof(name));
        ItemCount = itemCount;
        AssociatedItemCount = associatedItemCount;
    }

    /// <summary>Stable source identifier.</summary>
    public string Id { get; }

    /// <summary>Parent folder identifier, or null for a root.</summary>
    public string? ParentId { get; }

    /// <summary>Folder display name.</summary>
    public string Name { get; }

    /// <summary>Declared visible-item count when the source provides one.</summary>
    public int? ItemCount { get; }

    /// <summary>Declared folder-associated-item count when the source provides one.</summary>
    public int? AssociatedItemCount { get; }
}
