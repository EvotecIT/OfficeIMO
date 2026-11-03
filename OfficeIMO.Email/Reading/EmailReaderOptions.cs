namespace OfficeIMO.Email;

/// <summary>Immutable resource and preservation policy for <see cref="EmailDocumentReader"/>.</summary>
public sealed class EmailReaderOptions {
    /// <summary>Default bounded reader policy.</summary>
    public static EmailReaderOptions Default { get; } = new EmailReaderOptions();

    /// <summary>Creates reader options.</summary>
    public EmailReaderOptions(
        long maxInputBytes = 256L * 1024L * 1024L,
        int maxHeaderBytes = 1024 * 1024,
        int maxHeaderCount = 10000,
        int maxPartCount = 10000,
        int maxMimeDepth = 64,
        long maxAttachmentBytes = 128L * 1024L * 1024L,
        long maxTotalAttachmentBytes = 512L * 1024L * 1024L,
        int maxNestedMessageDepth = 16,
        bool includeAttachmentContent = true,
        bool preserveRawSource = false,
        int maxCompoundDirectoryEntries = 65536,
        int maxMapiPropertyCount = 100000,
        long maxDecodedPropertyBytes = 512L * 1024L * 1024L,
        int maxTnefAttributeCount = 100000,
        int maxAttachmentCount = 10000,
        bool includeEmbeddedMessages = true) {
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(maxInputBytes);
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(maxHeaderBytes);
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(maxHeaderCount);
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(maxPartCount);
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(maxMimeDepth);
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(maxAttachmentBytes);
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(maxTotalAttachmentBytes);
        ArgumentOutOfRangeException.ThrowIfNegative(maxNestedMessageDepth);
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(maxCompoundDirectoryEntries);
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(maxMapiPropertyCount);
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(maxDecodedPropertyBytes);
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(maxTnefAttributeCount);
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(maxAttachmentCount);

        MaxInputBytes = maxInputBytes;
        MaxHeaderBytes = maxHeaderBytes;
        MaxHeaderCount = maxHeaderCount;
        MaxPartCount = maxPartCount;
        MaxMimeDepth = maxMimeDepth;
        MaxAttachmentBytes = maxAttachmentBytes;
        MaxTotalAttachmentBytes = maxTotalAttachmentBytes;
        MaxNestedMessageDepth = maxNestedMessageDepth;
        IncludeAttachmentContent = includeAttachmentContent;
        PreserveRawSource = preserveRawSource;
        MaxCompoundDirectoryEntries = maxCompoundDirectoryEntries;
        MaxMapiPropertyCount = maxMapiPropertyCount;
        MaxDecodedPropertyBytes = maxDecodedPropertyBytes;
        MaxTnefAttributeCount = maxTnefAttributeCount;
        MaxAttachmentCount = maxAttachmentCount;
        IncludeEmbeddedMessages = includeEmbeddedMessages;
    }

    /// <summary>Maximum artifact size accepted by the reader.</summary>
    public long MaxInputBytes { get; }
    /// <summary>Maximum bytes allowed in one MIME header section.</summary>
    public int MaxHeaderBytes { get; }
    /// <summary>Maximum number of header fields in one entity.</summary>
    public int MaxHeaderCount { get; }
    /// <summary>Maximum MIME entity count.</summary>
    public int MaxPartCount { get; }
    /// <summary>Maximum nested MIME depth.</summary>
    public int MaxMimeDepth { get; }
    /// <summary>Maximum decoded bytes for one attachment.</summary>
    public long MaxAttachmentBytes { get; }
    /// <summary>Maximum aggregate decoded attachment bytes.</summary>
    public long MaxTotalAttachmentBytes { get; }
    /// <summary>Maximum embedded-message recursion depth.</summary>
    public int MaxNestedMessageDepth { get; }
    /// <summary>Whether decoded attachment content is retained.</summary>
    public bool IncludeAttachmentContent { get; }
    /// <summary>Whether embedded messages are parsed into nested documents. Attachment metadata remains available.</summary>
    public bool IncludeEmbeddedMessages { get; }
    /// <summary>Whether original artifact bytes are retained for explicit lossless writing.</summary>
    public bool PreserveRawSource { get; }
    /// <summary>Maximum CFB directory entries accepted while reading MSG.</summary>
    public int MaxCompoundDirectoryEntries { get; }
    /// <summary>Maximum aggregate MAPI properties across an artifact and embedded messages.</summary>
    public int MaxMapiPropertyCount { get; }
    /// <summary>Maximum aggregate decoded MSG/TNEF property bytes and MIME body/semantic-part bytes, checked before materialization.</summary>
    public long MaxDecodedPropertyBytes { get; }
    /// <summary>Maximum number of TNEF attributes.</summary>
    public int MaxTnefAttributeCount { get; }
    /// <summary>Maximum aggregate attachment count across the message and nested messages.</summary>
    public int MaxAttachmentCount { get; }
}
