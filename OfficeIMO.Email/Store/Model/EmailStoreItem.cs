using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

/// <summary>An email or typed Outlook item located within a store folder.</summary>
public sealed class EmailStoreItem {
    internal EmailStoreItem(string id, string folderId, EmailDocument document,
        bool isAssociated = false, bool isOrphaned = false,
        EmailStoreItemReadParts loadedParts = EmailStoreItemReadParts.All,
        EmailStoreFormat format = EmailStoreFormat.Unknown,
        EmailStoreItemSummary? summary = null) {
        Id = id;
        FolderId = folderId;
        Document = document ?? throw new ArgumentNullException(nameof(document));
        IsAssociated = isAssociated;
        IsOrphaned = isOrphaned;
        LoadedParts = loadedParts;
        ContentAvailability = EmailStoreItemContentAvailability.Create(
            format, document, loadedParts, summary);
    }

    /// <summary>Stable source identifier.</summary>
    public string Id { get; }

    /// <summary>Typed stable source identifier.</summary>
    public EmailStoreItemId Key => new EmailStoreItemId(Id);

    /// <summary>Containing folder identifier.</summary>
    public string FolderId { get; }

    /// <summary>Typed containing-folder identifier.</summary>
    public EmailStoreFolderId FolderKey => new EmailStoreFolderId(FolderId);

    /// <summary>Format-neutral projected item.</summary>
    public EmailDocument Document { get; }

    /// <summary>True for folder-associated information (FAI), such as views and folder settings.</summary>
    public bool IsAssociated { get; }

    /// <summary>True when the item was recovered from the NBT but is absent from the folder contents tables.</summary>
    public bool IsOrphaned { get; }

    /// <summary>
    /// Parts projected by the backend. A materializing backend can return more than the caller requested.
    /// </summary>
    public EmailStoreItemReadParts LoadedParts { get; }

    /// <summary>Local availability and partial-cache signals for the requested parts.</summary>
    public EmailStoreItemContentAvailability ContentAvailability { get; }
}
