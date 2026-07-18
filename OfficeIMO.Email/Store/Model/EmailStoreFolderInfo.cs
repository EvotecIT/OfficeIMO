namespace OfficeIMO.Email.Store;

/// <summary>Lightweight folder metadata exposed by an open email-store session.</summary>
public sealed class EmailStoreFolderInfo {
    private readonly List<MapiProperty> _mapiProperties;
    private MapiPropertyBag? _mapi;

    internal EmailStoreFolderInfo(string id, string? parentId, string name,
        int? itemCount = null, int? associatedItemCount = null,
        EmailStoreSpecialFolderKind specialFolderKind = EmailStoreSpecialFolderKind.Unknown,
        EmailStoreFolderClassificationSource classificationSource = EmailStoreFolderClassificationSource.None,
        string? containerClass = null, bool isSearchFolder = false,
        IEnumerable<MapiProperty>? mapiProperties = null) {
        Id = id ?? throw new ArgumentNullException(nameof(id));
        ParentId = parentId;
        Name = name ?? throw new ArgumentNullException(nameof(name));
        ItemCount = itemCount;
        AssociatedItemCount = associatedItemCount;
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

    /// <summary>Typed stable source identifier.</summary>
    public EmailStoreFolderId Key => new EmailStoreFolderId(Id);

    /// <summary>Parent folder identifier, or null for a root.</summary>
    public string? ParentId { get; }

    /// <summary>Typed parent folder identifier, or null for a root.</summary>
    public EmailStoreFolderId? ParentKey => ParentId == null ? (EmailStoreFolderId?)null : new EmailStoreFolderId(ParentId);

    /// <summary>Folder display name.</summary>
    public string Name { get; }

    /// <summary>Declared visible-item count when the source provides one.</summary>
    public int? ItemCount { get; }

    /// <summary>Declared folder-associated-item count when the source provides one.</summary>
    public int? AssociatedItemCount { get; }

    /// <summary>Well-known folder role, or <see cref="EmailStoreSpecialFolderKind.Unknown"/>.</summary>
    public EmailStoreSpecialFolderKind SpecialFolderKind { get; }

    /// <summary>Evidence used for the well-known classification.</summary>
    public EmailStoreFolderClassificationSource ClassificationSource { get; }

    /// <summary>MAPI container class, such as <c>IPF.Appointment</c>, when supplied by the source.</summary>
    public string? ContainerClass { get; }

    /// <summary>Whether the source identifies this as a search folder.</summary>
    public bool IsSearchFolder { get; }

    /// <summary>
    /// Detached folder MAPI-property snapshot. Changing this list does not mutate the open store.
    /// </summary>
    public IList<MapiProperty> MapiProperties => _mapiProperties;

    /// <summary>Typed access to the detached <see cref="MapiProperties"/> snapshot.</summary>
    public MapiPropertyBag Mapi => _mapi ?? (_mapi = new MapiPropertyBag(_mapiProperties));
}
