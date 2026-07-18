namespace OfficeIMO.Email.Store;

/// <summary>Safety and materialization limits for email-store readers.</summary>
public sealed class EmailStoreReaderOptions {
    /// <summary>Creates bounded reader options.</summary>
    public EmailStoreReaderOptions(
        long maxInputBytes = 1L * 1024 * 1024 * 1024 * 1024,
        int maxNodeCount = 25_000_000,
        int maxBTreeDepth = 32,
        int maxCachedBTreePages = 512,
        int maxFolderCount = 100_000,
        int maxItemCount = 1_000_000,
        int maxPropertiesPerItem = 16_384,
        long maxDecodedPropertyBytesPerItem = 128L * 1024 * 1024,
        int maxAttachmentsPerItem = 10_000,
        long maxAttachmentBytes = 512L * 1024 * 1024,
        long maxTotalAttachmentBytes = 4L * 1024 * 1024 * 1024,
        bool retainAttachmentContent = true,
        string? pstPassword = null,
        Encoding? pstPasswordEncoding = null,
        bool includeAssociatedItems = false,
        bool includeOrphanedItems = false,
        int maxNestedMessageDepth = 16,
        int maxArchiveEntries = 500_000,
        long maxArchiveEntryBytes = 512L * 1024 * 1024,
        long maxArchiveDecodedBytes = 8L * 1024 * 1024 * 1024,
        long maxXmlCharactersPerItem = 64L * 1024 * 1024,
        long maxMessageBytes = 256L * 1024 * 1024,
        int maxDirectoryDepth = 64,
        int maxDirectoryFileCount = 1_000_000,
        long maxDecodedTableBytes = 8L * 1024 * 1024 * 1024) {
        MaxInputBytes = Positive(maxInputBytes, nameof(maxInputBytes));
        MaxNodeCount = Positive(maxNodeCount, nameof(maxNodeCount));
        MaxBTreeDepth = Positive(maxBTreeDepth, nameof(maxBTreeDepth));
        MaxCachedBTreePages = Positive(maxCachedBTreePages, nameof(maxCachedBTreePages));
        MaxFolderCount = Positive(maxFolderCount, nameof(maxFolderCount));
        MaxItemCount = Positive(maxItemCount, nameof(maxItemCount));
        MaxPropertiesPerItem = Positive(maxPropertiesPerItem, nameof(maxPropertiesPerItem));
        MaxDecodedPropertyBytesPerItem = Positive(maxDecodedPropertyBytesPerItem, nameof(maxDecodedPropertyBytesPerItem));
        MaxAttachmentsPerItem = Positive(maxAttachmentsPerItem, nameof(maxAttachmentsPerItem));
        MaxAttachmentBytes = Positive(maxAttachmentBytes, nameof(maxAttachmentBytes));
        MaxTotalAttachmentBytes = Positive(maxTotalAttachmentBytes, nameof(maxTotalAttachmentBytes));
        RetainAttachmentContent = retainAttachmentContent;
        PstPassword = pstPassword;
        PstPasswordEncoding = pstPasswordEncoding ?? Encoding.ASCII;
        IncludeAssociatedItems = includeAssociatedItems;
        IncludeOrphanedItems = includeOrphanedItems;
        if (maxNestedMessageDepth < 0) throw new ArgumentOutOfRangeException(nameof(maxNestedMessageDepth));
        MaxNestedMessageDepth = maxNestedMessageDepth;
        MaxArchiveEntries = Positive(maxArchiveEntries, nameof(maxArchiveEntries));
        MaxArchiveEntryBytes = Positive(maxArchiveEntryBytes, nameof(maxArchiveEntryBytes));
        MaxArchiveDecodedBytes = Positive(maxArchiveDecodedBytes, nameof(maxArchiveDecodedBytes));
        MaxXmlCharactersPerItem = Positive(maxXmlCharactersPerItem, nameof(maxXmlCharactersPerItem));
        MaxMessageBytes = Positive(maxMessageBytes, nameof(maxMessageBytes));
        MaxDirectoryDepth = Positive(maxDirectoryDepth, nameof(maxDirectoryDepth));
        MaxDirectoryFileCount = Positive(maxDirectoryFileCount, nameof(maxDirectoryFileCount));
        MaxDecodedTableBytes = Positive(maxDecodedTableBytes, nameof(maxDecodedTableBytes));
    }

    /// <summary>Default bounded options.</summary>
    public static EmailStoreReaderOptions Default { get; } = new EmailStoreReaderOptions();

    /// <summary>Maximum seekable source length.</summary>
    public long MaxInputBytes { get; }
    /// <summary>Maximum NDB nodes and blocks visited.</summary>
    public int MaxNodeCount { get; }
    /// <summary>Maximum tree traversal depth.</summary>
    public int MaxBTreeDepth { get; }
    /// <summary>Maximum PST/OST B-tree pages retained by the bounded random-access cache.</summary>
    public int MaxCachedBTreePages { get; }
    /// <summary>Maximum folders materialized.</summary>
    public int MaxFolderCount { get; }
    /// <summary>Maximum items materialized.</summary>
    public int MaxItemCount { get; }
    /// <summary>Maximum MAPI properties decoded per item.</summary>
    public int MaxPropertiesPerItem { get; }
    /// <summary>Maximum decoded property bytes per item.</summary>
    public long MaxDecodedPropertyBytesPerItem { get; }
    /// <summary>Maximum attachments per item.</summary>
    public int MaxAttachmentsPerItem { get; }
    /// <summary>Maximum decoded bytes in one attachment.</summary>
    public long MaxAttachmentBytes { get; }
    /// <summary>Maximum decoded attachment bytes across the read.</summary>
    public long MaxTotalAttachmentBytes { get; }
    /// <summary>Whether attachment payloads are retained in memory.</summary>
    public bool RetainAttachmentContent { get; }
    /// <summary>Password to validate when PidTagPstPassword is nonzero. The value is never logged or retained by results.</summary>
    public string? PstPassword { get; }
    /// <summary>Byte encoding used for the legacy PST password checksum. Defaults to ASCII.</summary>
    public Encoding PstPasswordEncoding { get; }
    /// <summary>Whether folder-associated information items are materialized separately from visible items.</summary>
    public bool IncludeAssociatedItems { get; }
    /// <summary>Whether NBT item nodes absent from folder contents tables are recovered using their parent links.</summary>
    public bool IncludeOrphanedItems { get; }
    /// <summary>Maximum embedded-message recursion depth. Zero preserves the attachment without projecting its item.</summary>
    public int MaxNestedMessageDepth { get; }
    /// <summary>Maximum entries accepted from a compressed email-store archive.</summary>
    public int MaxArchiveEntries { get; }
    /// <summary>Maximum decoded size declared by one compressed archive entry.</summary>
    public long MaxArchiveEntryBytes { get; }
    /// <summary>Maximum total decoded size declared by all compressed archive entries.</summary>
    public long MaxArchiveDecodedBytes { get; }
    /// <summary>Maximum XML characters parsed from one archive item.</summary>
    public long MaxXmlCharactersPerItem { get; }
    /// <summary>Maximum RFC 5322/MIME message bytes accepted from one store item.</summary>
    public long MaxMessageBytes { get; }
    /// <summary>Maximum directory depth traversed by mailbox-directory sessions.</summary>
    public int MaxDirectoryDepth { get; }
    /// <summary>Maximum EML, EMLX, and Maildir files indexed by one mailbox-directory session.</summary>
    public int MaxDirectoryFileCount { get; }
    /// <summary>Maximum decoded bytes traversed from one PST/OST table data tree.</summary>
    public long MaxDecodedTableBytes { get; }

    private static int Positive(int value, string name) {
        if (value <= 0) throw new ArgumentOutOfRangeException(name);
        return value;
    }

    private static long Positive(long value, string name) {
        if (value <= 0) throw new ArgumentOutOfRangeException(name);
        return value;
    }
}
