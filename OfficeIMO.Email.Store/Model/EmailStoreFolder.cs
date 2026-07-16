namespace OfficeIMO.Email.Store;

/// <summary>A folder in an email store.</summary>
public sealed class EmailStoreFolder {
    private readonly List<EmailStoreMessage> _messages = new List<EmailStoreMessage>();

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

    internal IList<EmailStoreMessage> MutableMessages => _messages;
}
