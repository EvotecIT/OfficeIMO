namespace OfficeIMO.Email.Store;

/// <summary>Lightweight folder metadata exposed by an open email-store session.</summary>
public sealed class EmailStoreFolderInfo {
    internal EmailStoreFolderInfo(string id, string? parentId, string name,
        int? itemCount = null, int? associatedItemCount = null,
        EmailStoreSpecialFolderKind specialFolderKind = EmailStoreSpecialFolderKind.Unknown,
        EmailStoreFolderClassificationSource classificationSource = EmailStoreFolderClassificationSource.None,
        string? containerClass = null, bool isSearchFolder = false) {
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

    /// <summary>Well-known folder role, or <see cref="EmailStoreSpecialFolderKind.Unknown"/>.</summary>
    public EmailStoreSpecialFolderKind SpecialFolderKind { get; }

    /// <summary>Evidence used for the well-known classification.</summary>
    public EmailStoreFolderClassificationSource ClassificationSource { get; }

    /// <summary>MAPI container class, such as <c>IPF.Appointment</c>, when supplied by the source.</summary>
    public string? ContainerClass { get; }

    /// <summary>Whether the source identifies this as a search folder.</summary>
    public bool IsSearchFolder { get; }
}
