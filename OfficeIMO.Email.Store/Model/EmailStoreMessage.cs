using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

/// <summary>An email or Outlook item located within a store folder.</summary>
public sealed class EmailStoreMessage {
    internal EmailStoreMessage(string id, string folderId, EmailDocument document,
        bool isAssociated = false, bool isOrphaned = false) {
        Id = id;
        FolderId = folderId;
        Document = document ?? throw new ArgumentNullException(nameof(document));
        IsAssociated = isAssociated;
        IsOrphaned = isOrphaned;
    }

    /// <summary>Stable source identifier.</summary>
    public string Id { get; }

    /// <summary>Containing folder identifier.</summary>
    public string FolderId { get; }

    /// <summary>Format-neutral projected item.</summary>
    public EmailDocument Document { get; }

    /// <summary>True for folder-associated information (FAI), such as views and folder settings.</summary>
    public bool IsAssociated { get; }

    /// <summary>True when the item was recovered from the NBT but is absent from the folder contents tables.</summary>
    public bool IsOrphaned { get; }
}
