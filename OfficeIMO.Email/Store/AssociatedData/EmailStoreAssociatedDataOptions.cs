namespace OfficeIMO.Email.Store;

/// <summary>Bounds and scope for reading folder-associated information (FAI).</summary>
public sealed class EmailStoreAssociatedDataOptions {
    /// <summary>Creates immutable associated-data read options.</summary>
    public EmailStoreAssociatedDataOptions(
        EmailStoreFolderId? folderId = null,
        bool includeDescendants = false,
        int maxItems = 10_000,
        long maxDecodedPropertyBytesPerItem = 16 * 1024 * 1024,
        int maxXmlBytes = 4 * 1024 * 1024,
        bool continueOnError = true) {
        if (folderId.HasValue && folderId.Value.IsEmpty) {
            throw new ArgumentException("The folder identifier cannot be empty.", nameof(folderId));
        }
        if (maxItems <= 0) throw new ArgumentOutOfRangeException(nameof(maxItems));
        if (maxDecodedPropertyBytesPerItem <= 0) {
            throw new ArgumentOutOfRangeException(nameof(maxDecodedPropertyBytesPerItem));
        }
        if (maxXmlBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maxXmlBytes));
        FolderId = folderId;
        IncludeDescendants = includeDescendants;
        MaxItems = maxItems;
        MaxDecodedPropertyBytesPerItem = maxDecodedPropertyBytesPerItem;
        MaxXmlBytes = maxXmlBytes;
        ContinueOnError = continueOnError;
    }

    /// <summary>Optional folder scope; null includes the entire Store.</summary>
    public EmailStoreFolderId? FolderId { get; }

    /// <summary>Whether descendants of <see cref="FolderId"/> are included.</summary>
    public bool IncludeDescendants { get; }

    /// <summary>Maximum associated messages read.</summary>
    public int MaxItems { get; }

    /// <summary>Maximum decoded root-property bytes for any one associated message.</summary>
    public long MaxDecodedPropertyBytesPerItem { get; }

    /// <summary>Maximum bytes accepted for any XML configuration stream.</summary>
    public int MaxXmlBytes { get; }

    /// <summary>Whether corrupt individual associated messages are reported and skipped.</summary>
    public bool ContinueOnError { get; }
}
