namespace OfficeIMO.Email.Store;

/// <summary>Lightweight stable reference to an item that can be read explicitly from a session.</summary>
public sealed class EmailStoreItemReference {
    internal EmailStoreItemReference(string id, string folderId, bool isAssociated, bool isOrphaned,
        EmailStoreItemSummary? summary = null) {
        Id = id ?? throw new ArgumentNullException(nameof(id));
        FolderId = folderId ?? throw new ArgumentNullException(nameof(folderId));
        IsAssociated = isAssociated;
        IsOrphaned = isOrphaned;
        Summary = summary;
    }

    /// <summary>Stable source identifier.</summary>
    public string Id { get; }

    /// <summary>Typed stable source identifier.</summary>
    public EmailStoreItemId Key => new EmailStoreItemId(Id);

    /// <summary>Containing folder identifier.</summary>
    public string FolderId { get; }

    /// <summary>Typed containing-folder identifier.</summary>
    public EmailStoreFolderId FolderKey => new EmailStoreFolderId(FolderId);

    /// <summary>True for folder-associated information such as views and folder settings.</summary>
    public bool IsAssociated { get; }

    /// <summary>True when the source index exposes the item but its folder contents table does not.</summary>
    public bool IsOrphaned { get; }

    /// <summary>Summary projected from a source contents table when it was available without an item read.</summary>
    public EmailStoreItemSummary? Summary { get; }
}
