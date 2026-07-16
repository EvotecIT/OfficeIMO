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
        bool retainAttachmentContent = true) {
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

    private static int Positive(int value, string name) {
        if (value <= 0) throw new ArgumentOutOfRangeException(name);
        return value;
    }

    private static long Positive(long value, string name) {
        if (value <= 0) throw new ArgumentOutOfRangeException(name);
        return value;
    }
}
