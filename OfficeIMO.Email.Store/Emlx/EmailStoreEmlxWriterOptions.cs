using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

/// <summary>Deterministic policy for writing Apple Mail EMLX artifacts.</summary>
public sealed class EmailStoreEmlxWriterOptions {
    /// <summary>Default EMLX writer policy.</summary>
    public static EmailStoreEmlxWriterOptions Default { get; } = new EmailStoreEmlxWriterOptions();

    /// <summary>Creates an EMLX writer policy.</summary>
    public EmailStoreEmlxWriterOptions(EmailWriterOptions? messageOptions = null,
        bool includeMetadata = true, long maxOutputBytes = 512L * 1024L * 1024L) {
        if (maxOutputBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maxOutputBytes));
        MessageOptions = messageOptions ?? EmailWriterOptions.Default;
        IncludeMetadata = includeMetadata;
        MaxOutputBytes = maxOutputBytes;
    }

    /// <summary>EML serialization policy.</summary>
    public EmailWriterOptions MessageOptions { get; }
    /// <summary>Whether an Apple property-list metadata trailer is emitted.</summary>
    public bool IncludeMetadata { get; }
    /// <summary>Maximum complete EMLX artifact bytes.</summary>
    public long MaxOutputBytes { get; }
}
