namespace OfficeIMO.Email.Store;

/// <summary>A folder in an email store.</summary>
public sealed class EmailStoreFolder {
    private readonly List<EmailStoreItem> _items = new List<EmailStoreItem>();
    private readonly List<EmailStoreItem> _associatedItems = new List<EmailStoreItem>();
    private readonly List<MapiProperty> _mapiProperties;
    private MapiPropertyBag? _mapi;

    internal EmailStoreFolder(string id, string? parentId, string name,
        EmailStoreSpecialFolderKind specialFolderKind = EmailStoreSpecialFolderKind.Unknown,
        EmailStoreFolderClassificationSource classificationSource = EmailStoreFolderClassificationSource.None,
        string? containerClass = null, bool isSearchFolder = false,
        IEnumerable<MapiProperty>? mapiProperties = null) {
        Id = id;
        ParentId = parentId;
        Name = name;
        if (specialFolderKind == EmailStoreSpecialFolderKind.Unknown) {
            specialFolderKind = EmailStoreSpecialFolderClassifier.FromDisplayName(name);
            if (specialFolderKind != EmailStoreSpecialFolderKind.Unknown) {
                classificationSource = EmailStoreFolderClassificationSource.DisplayName;
            }
        }
        SpecialFolderKind = specialFolderKind;
        ClassificationSource = classificationSource;
        ContainerClass = containerClass;
        IsSearchFolder = isSearchFolder;
        _mapiProperties = mapiProperties?.Select(MapiPropertySnapshot.Clone).ToList() ?? new List<MapiProperty>();
    }

    /// <summary>Stable source identifier.</summary>
    public string Id { get; }

    /// <summary>Parent folder identifier, or null for a root.</summary>
    public string? ParentId { get; }

    /// <summary>Display name.</summary>
    public string Name { get; internal set; }

    /// <summary>Well-known folder role, or <see cref="EmailStoreSpecialFolderKind.Unknown"/>.</summary>
    public EmailStoreSpecialFolderKind SpecialFolderKind { get; }

    /// <summary>Evidence used for the well-known classification.</summary>
    public EmailStoreFolderClassificationSource ClassificationSource { get; }

    /// <summary>MAPI container class when supplied by the source.</summary>
    public string? ContainerClass { get; }

    /// <summary>Whether the source identifies this as a search folder.</summary>
    public bool IsSearchFolder { get; }

    /// <summary>Detached folder MAPI-property snapshot.</summary>
    public IList<MapiProperty> MapiProperties => _mapiProperties;

    /// <summary>Typed access to the detached <see cref="MapiProperties"/> snapshot.</summary>
    public MapiPropertyBag Mapi => _mapi ?? (_mapi = new MapiPropertyBag(_mapiProperties));

    /// <summary>Items directly contained in this folder.</summary>
    public IReadOnlyList<EmailStoreItem> Items => _items;

    /// <summary>Folder-associated information items when explicitly requested by reader options.</summary>
    public IReadOnlyList<EmailStoreItem> AssociatedItems => _associatedItems;

    internal IList<EmailStoreItem> MutableItems => _items;
    internal IList<EmailStoreItem> MutableAssociatedItems => _associatedItems;
}
