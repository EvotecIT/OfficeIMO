namespace OfficeIMO.Email.Store;

/// <summary>Materialized mailbox hierarchy and its projected email and Outlook items.</summary>
public sealed class EmailStore {
    private readonly List<EmailStoreFolder> _folders = new List<EmailStoreFolder>();

    /// <summary>Detected source format.</summary>
    public EmailStoreFormat Format { get; internal set; }

    /// <summary>Display name declared by the store, when available.</summary>
    public string? DisplayName { get; internal set; }

    /// <summary>All folders in stable source order. Parent identifiers preserve the original hierarchy.</summary>
    public IReadOnlyList<EmailStoreFolder> Folders => _folders;

    /// <summary>Root-level folders.</summary>
    public IEnumerable<EmailStoreFolder> RootFolders => _folders.Where(folder => folder.ParentId == null);

    /// <summary>Total projected item count.</summary>
    public int ItemCount => _folders.Sum(folder => folder.Items.Count);

    internal IList<EmailStoreFolder> MutableFolders => _folders;
}
