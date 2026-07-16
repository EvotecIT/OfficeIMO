namespace OfficeIMO.Email.Store;

/// <summary>Controls which parts are decoded when one selected store item is read.</summary>
public sealed class EmailStoreItemReadOptions {
    /// <summary>Creates selective item options. Dependent parts are included automatically.</summary>
    public EmailStoreItemReadOptions(EmailStoreItemReadParts parts = EmailStoreItemReadParts.All,
        long? maxDecodedPropertyBytes = null,
        bool preferStreamingAttachmentContent = false) {
        const EmailStoreItemReadParts known = EmailStoreItemReadParts.All;
        if ((parts & ~known) != 0) throw new ArgumentOutOfRangeException(nameof(parts));
        if (maxDecodedPropertyBytes.HasValue && maxDecodedPropertyBytes.Value <= 0) {
            throw new ArgumentOutOfRangeException(nameof(maxDecodedPropertyBytes));
        }
        Parts = Normalize(parts);
        MaxDecodedPropertyBytes = maxDecodedPropertyBytes;
        PreferStreamingAttachmentContent = preferStreamingAttachmentContent;
    }

    /// <summary>Default full-item projection.</summary>
    public static EmailStoreItemReadOptions Default { get; } = new EmailStoreItemReadOptions();

    /// <summary>Normalized parts requested from the store backend.</summary>
    public EmailStoreItemReadParts Parts { get; }

    /// <summary>
    /// Optional per-read cap for decoded property bytes. Backends may narrow this further and never widen
    /// <see cref="EmailStoreReaderOptions.MaxDecodedPropertyBytesPerItem"/>.
    /// </summary>
    public long? MaxDecodedPropertyBytes { get; }

    /// <summary>
    /// Whether a capable session should expose attachment payloads through reopenable streams even when the
    /// store's compatibility options normally retain byte arrays.
    /// </summary>
    public bool PreferStreamingAttachmentContent { get; }

    /// <summary>Returns true when a part is requested.</summary>
    public bool Includes(EmailStoreItemReadParts part) => (Parts & part) == part;

    private static EmailStoreItemReadParts Normalize(EmailStoreItemReadParts parts) {
        if ((parts & EmailStoreItemReadParts.AttachmentContent) != 0) {
            parts |= EmailStoreItemReadParts.AttachmentMetadata;
        }
        if ((parts & EmailStoreItemReadParts.EmbeddedItems) != 0) {
            parts |= EmailStoreItemReadParts.AttachmentMetadata;
        }
        if (parts != EmailStoreItemReadParts.None) parts |= EmailStoreItemReadParts.Metadata;
        return parts;
    }
}
