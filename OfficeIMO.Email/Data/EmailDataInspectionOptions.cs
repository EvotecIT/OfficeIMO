using OfficeIMO.Email.AddressBook;
using OfficeIMO.Email.Store;

namespace OfficeIMO.Email.Data;

/// <summary>Owner-level read limits and finite metadata samples for a local email-data inspection.</summary>
public sealed class EmailDataInspectionOptions {
    /// <summary>Creates a bounded report policy. An explicit open policy retains its own owner-specific limits.</summary>
    public EmailDataInspectionOptions(EmailDataOpenOptions? openOptions = null, int maxSamples = 64,
        int maxPreviewCharacters = 256, int maxHeadersInspected = 10000) {
        if (maxSamples <= 0 || maxSamples > 1000) throw new ArgumentOutOfRangeException(nameof(maxSamples));
        if (maxPreviewCharacters <= 0 || maxPreviewCharacters > 1024) throw new ArgumentOutOfRangeException(nameof(maxPreviewCharacters));
        if (maxHeadersInspected <= 0 || maxHeadersInspected > 100000) throw new ArgumentOutOfRangeException(nameof(maxHeadersInspected));
        MaxSamples = maxSamples; MaxPreviewCharacters = maxPreviewCharacters; MaxHeadersInspected = maxHeadersInspected;
        OpenOptions = openOptions ?? new EmailDataOpenOptions(
            email: new EmailReaderOptions(maxInputBytes: 16L * 1024 * 1024,
                maxDecodedPropertyBytes: 16L * 1024 * 1024, includeAttachmentContent: false,
                includeEmbeddedMessages: false),
            contentLines: new ContentLineReaderOptions(maxComponents: 10000, maxProperties: 100000),
            store: new EmailStoreReaderOptions(maxInputBytes: 4L * 1024 * 1024 * 1024,
                maxNodeCount: 1000000, maxFolderCount: 1000, maxItemCount: 10000,
                maxDecodedPropertyBytesPerItem: 16L * 1024 * 1024, retainAttachmentContent: false,
                maxArchiveEntries: 10000, maxArchiveEntryBytes: 16L * 1024 * 1024,
                maxArchiveDecodedBytes: 256L * 1024 * 1024, maxXmlCharactersPerItem: 16L * 1024 * 1024,
                maxMessageBytes: 16L * 1024 * 1024, maxDirectoryFileCount: 10000,
                maxDecodedTableBytes: 256L * 1024 * 1024),
            addressBook: new OfflineAddressBookReaderOptions(maxInputBytes: 64L * 1024 * 1024,
                maxDiscoveredFiles: 64, maxDeclaredEntries: 1000000, retainRawPropertyBytes: false,
                maxDirectoryEntries: 10000), useStreamingEmailReader: true);
    }
    /// <summary>Policy applied before each existing owner opens the artifact.</summary>
    public EmailDataOpenOptions OpenOptions { get; }
    /// <summary>Maximum rows retained in each sampled metadata collection.</summary>
    public int MaxSamples { get; }
    /// <summary>Maximum UTF-16 units in each metadata preview, without splitting surrogate pairs.</summary>
    public int MaxPreviewCharacters { get; }
    /// <summary>Maximum envelope headers scanned for transport signature names.</summary>
    public int MaxHeadersInspected { get; }
}
