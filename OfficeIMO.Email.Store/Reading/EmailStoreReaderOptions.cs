namespace OfficeIMO.Email.Store;

/// <summary>Safety and materialization limits for email-store readers.</summary>
public sealed class EmailStoreReaderOptions {
    /// <summary>Creates bounded reader options.</summary>
    public EmailStoreReaderOptions(
        long maxInputBytes = 8L * 1024 * 1024 * 1024,
        int maxNodeCount = 2_000_000,
        int maxBTreeDepth = 32,
        int maxFolderCount = 100_000,
        int maxMessageCount = 1_000_000,
        int maxPropertiesPerItem = 16_384,
        long maxDecodedPropertyBytesPerItem = 128L * 1024 * 1024,
        int maxAttachmentsPerMessage = 10_000,
        long maxAttachmentBytes = 512L * 1024 * 1024,
        long maxTotalAttachmentBytes = 4L * 1024 * 1024 * 1024,
        bool retainAttachmentContent = true,
        string? pstPassword = null,
        Encoding? pstPasswordEncoding = null,
        bool includeAssociatedMessages = false,
        bool includeOrphanedMessages = false,
        int maxNestedMessageDepth = 16) {
        MaxInputBytes = Positive(maxInputBytes, nameof(maxInputBytes));
        MaxNodeCount = Positive(maxNodeCount, nameof(maxNodeCount));
        MaxBTreeDepth = Positive(maxBTreeDepth, nameof(maxBTreeDepth));
        MaxFolderCount = Positive(maxFolderCount, nameof(maxFolderCount));
        MaxMessageCount = Positive(maxMessageCount, nameof(maxMessageCount));
        MaxPropertiesPerItem = Positive(maxPropertiesPerItem, nameof(maxPropertiesPerItem));
        MaxDecodedPropertyBytesPerItem = Positive(maxDecodedPropertyBytesPerItem, nameof(maxDecodedPropertyBytesPerItem));
        MaxAttachmentsPerMessage = Positive(maxAttachmentsPerMessage, nameof(maxAttachmentsPerMessage));
        MaxAttachmentBytes = Positive(maxAttachmentBytes, nameof(maxAttachmentBytes));
        MaxTotalAttachmentBytes = Positive(maxTotalAttachmentBytes, nameof(maxTotalAttachmentBytes));
        RetainAttachmentContent = retainAttachmentContent;
        PstPassword = pstPassword;
        PstPasswordEncoding = pstPasswordEncoding ?? Encoding.ASCII;
        IncludeAssociatedMessages = includeAssociatedMessages;
        IncludeOrphanedMessages = includeOrphanedMessages;
        if (maxNestedMessageDepth < 0) throw new ArgumentOutOfRangeException(nameof(maxNestedMessageDepth));
        MaxNestedMessageDepth = maxNestedMessageDepth;
    }

    /// <summary>Default bounded options.</summary>
    public static EmailStoreReaderOptions Default { get; } = new EmailStoreReaderOptions();

    /// <summary>Maximum seekable source length.</summary>
    public long MaxInputBytes { get; }
    /// <summary>Maximum NDB nodes and blocks visited.</summary>
    public int MaxNodeCount { get; }
    /// <summary>Maximum tree traversal depth.</summary>
    public int MaxBTreeDepth { get; }
    /// <summary>Maximum folders materialized.</summary>
    public int MaxFolderCount { get; }
    /// <summary>Maximum messages materialized.</summary>
    public int MaxMessageCount { get; }
    /// <summary>Maximum MAPI properties decoded per item.</summary>
    public int MaxPropertiesPerItem { get; }
    /// <summary>Maximum decoded property bytes per item.</summary>
    public long MaxDecodedPropertyBytesPerItem { get; }
    /// <summary>Maximum attachments per message.</summary>
    public int MaxAttachmentsPerMessage { get; }
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
    /// <summary>Whether folder-associated information items are materialized separately from visible messages.</summary>
    public bool IncludeAssociatedMessages { get; }
    /// <summary>Whether NBT message nodes absent from folder contents tables are recovered using their parent links.</summary>
    public bool IncludeOrphanedMessages { get; }
    /// <summary>Maximum embedded-message recursion depth. Zero preserves the attachment without projecting its item.</summary>
    public int MaxNestedMessageDepth { get; }

    private static int Positive(int value, string name) {
        if (value <= 0) throw new ArgumentOutOfRangeException(name);
        return value;
    }

    private static long Positive(long value, string name) {
        if (value <= 0) throw new ArgumentOutOfRangeException(name);
        return value;
    }
}
