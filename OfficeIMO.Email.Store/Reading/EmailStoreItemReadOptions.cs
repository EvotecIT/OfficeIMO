namespace OfficeIMO.Email.Store;

/// <summary>Controls which parts are decoded when one selected store item is read.</summary>
public sealed class EmailStoreItemReadOptions {
    /// <summary>Creates selective item options. Dependent parts are included automatically.</summary>
    public EmailStoreItemReadOptions(EmailStoreItemReadParts parts = EmailStoreItemReadParts.All) {
        const EmailStoreItemReadParts known = EmailStoreItemReadParts.All;
        if ((parts & ~known) != 0) throw new ArgumentOutOfRangeException(nameof(parts));
        Parts = Normalize(parts);
    }

    /// <summary>Default full-item projection.</summary>
    public static EmailStoreItemReadOptions Default { get; } = new EmailStoreItemReadOptions();

    /// <summary>Normalized parts requested from the store backend.</summary>
    public EmailStoreItemReadParts Parts { get; }

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
