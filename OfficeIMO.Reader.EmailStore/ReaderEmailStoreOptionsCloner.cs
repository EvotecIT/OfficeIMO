using OfficeIMO.Email.Store;
using System.Text;

namespace OfficeIMO.Reader.EmailStore;

internal static class ReaderEmailStoreOptionsCloner {
    internal static ReaderEmailStoreOptions CloneOrDefault(ReaderEmailStoreOptions? options) {
        return new ReaderEmailStoreOptions {
            StoreOptions = CloneStoreOptions(options?.StoreOptions ?? EmailStoreReaderOptions.Default)
        };
    }

    internal static EmailStoreReaderOptions CreateEffective(
        ReaderEmailStoreOptions options,
        ReaderOptions readerOptions) {
        if (options == null) throw new ArgumentNullException(nameof(options));
        if (readerOptions == null) throw new ArgumentNullException(nameof(readerOptions));
        EmailStoreReaderOptions source = options.StoreOptions ?? EmailStoreReaderOptions.Default;
        long maxInputBytes = source.MaxInputBytes;
        if (readerOptions.MaxInputBytes.HasValue && readerOptions.MaxInputBytes.Value > 0) {
            maxInputBytes = Math.Min(maxInputBytes, readerOptions.MaxInputBytes.Value);
        }
        return CloneStoreOptions(source, maxInputBytes);
    }

    private static EmailStoreReaderOptions CloneStoreOptions(
        EmailStoreReaderOptions source,
        long? maxInputBytes = null) {
        Encoding passwordEncoding = (Encoding)source.PstPasswordEncoding.Clone();
        return new EmailStoreReaderOptions(
            maxInputBytes: maxInputBytes ?? source.MaxInputBytes,
            maxNodeCount: source.MaxNodeCount,
            maxBTreeDepth: source.MaxBTreeDepth,
            maxFolderCount: source.MaxFolderCount,
            maxItemCount: source.MaxItemCount,
            maxPropertiesPerItem: source.MaxPropertiesPerItem,
            maxDecodedPropertyBytesPerItem: source.MaxDecodedPropertyBytesPerItem,
            maxAttachmentsPerItem: source.MaxAttachmentsPerItem,
            maxAttachmentBytes: source.MaxAttachmentBytes,
            maxTotalAttachmentBytes: source.MaxTotalAttachmentBytes,
            retainAttachmentContent: source.RetainAttachmentContent,
            pstPassword: source.PstPassword,
            pstPasswordEncoding: passwordEncoding,
            includeAssociatedItems: source.IncludeAssociatedItems,
            includeOrphanedItems: source.IncludeOrphanedItems,
            maxNestedMessageDepth: source.MaxNestedMessageDepth,
            maxArchiveEntries: source.MaxArchiveEntries,
            maxArchiveEntryBytes: source.MaxArchiveEntryBytes,
            maxArchiveDecodedBytes: source.MaxArchiveDecodedBytes,
            maxXmlCharactersPerItem: source.MaxXmlCharactersPerItem,
            maxMessageBytes: source.MaxMessageBytes);
    }
}
