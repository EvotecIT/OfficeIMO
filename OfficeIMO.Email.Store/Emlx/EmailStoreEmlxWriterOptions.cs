using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

/// <summary>Deterministic policy for writing Apple Mail EMLX artifacts.</summary>
public sealed class EmailStoreEmlxWriterOptions {
    /// <summary>Default EMLX writer policy.</summary>
    public static EmailStoreEmlxWriterOptions Default { get; } = new EmailStoreEmlxWriterOptions();

    /// <summary>Creates an EMLX writer policy.</summary>
    public EmailStoreEmlxWriterOptions(EmailWriterOptions? messageOptions = null,
        bool includeMetadata = true, long maxOutputBytes = 512L * 1024L * 1024L,
        int maxMetadataDepth = 32) {
        if (maxOutputBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maxOutputBytes));
        if (maxMetadataDepth <= 0) throw new ArgumentOutOfRangeException(nameof(maxMetadataDepth));
        MessageOptions = messageOptions ?? EmailWriterOptions.Default;
        IncludeMetadata = includeMetadata;
        MaxOutputBytes = maxOutputBytes;
        MaxMetadataDepth = maxMetadataDepth;
    }

    /// <summary>EML serialization policy.</summary>
    public EmailWriterOptions MessageOptions { get; }
    /// <summary>Whether an Apple property-list metadata trailer is emitted.</summary>
    public bool IncludeMetadata { get; }
    /// <summary>Maximum complete EMLX artifact bytes.</summary>
    public long MaxOutputBytes { get; }
    /// <summary>
    /// Maximum retained property-list value depth, measured from the root dictionary at depth zero.
    /// The default matches <see cref="EmailStoreReaderOptions.MaxBTreeDepth"/>.
    /// </summary>
    public int MaxMetadataDepth { get; }
}
