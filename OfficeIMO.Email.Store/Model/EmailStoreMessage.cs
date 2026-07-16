using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

/// <summary>An email or Outlook item located within a store folder.</summary>
public sealed class EmailStoreMessage {
    internal EmailStoreMessage(string id, string folderId, EmailDocument document) {
        Id = id;
        FolderId = folderId;
        Document = document ?? throw new ArgumentNullException(nameof(document));
    }

    /// <summary>Stable source identifier.</summary>
    public string Id { get; }

    /// <summary>Containing folder identifier.</summary>
    public string FolderId { get; }

    /// <summary>Format-neutral projected item.</summary>
    public EmailDocument Document { get; }
}
