namespace OfficeIMO.Email.Store;

/// <summary>A folder in an email store.</summary>
public sealed class EmailStoreFolder {
    private readonly List<EmailStoreMessage> _messages = new List<EmailStoreMessage>();
    private readonly List<EmailStoreMessage> _associatedMessages = new List<EmailStoreMessage>();

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

    /// <summary>Messages directly contained in this folder.</summary>
    public IReadOnlyList<EmailStoreMessage> Messages => _messages;

    /// <summary>Folder-associated information items when explicitly requested by reader options.</summary>
    public IReadOnlyList<EmailStoreMessage> AssociatedMessages => _associatedMessages;

    internal IList<EmailStoreMessage> MutableMessages => _messages;
    internal IList<EmailStoreMessage> MutableAssociatedMessages => _associatedMessages;
}
