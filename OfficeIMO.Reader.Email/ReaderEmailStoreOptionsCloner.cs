using OfficeIMO.Email.Store;
using System.Text;

namespace OfficeIMO.Reader.Email;

internal static class ReaderEmailStoreOptionsCloner {
    internal static ReaderEmailStoreOptions CloneOrDefault(ReaderEmailStoreOptions? options) {
        return new ReaderEmailStoreOptions {
            StoreOptions = CloneStoreOptions(options?.StoreOptions ?? EmailStoreReaderOptions.Default),
            Query = CloneQuery(options?.Query),
            ItemReadOptions = CloneItemReadOptions(options?.ItemReadOptions),
            StreamAttachmentContent = options?.StreamAttachmentContent ?? true,
            MaxItems = options?.MaxItems ?? 1_000,
            ContinueOnItemError = options?.ContinueOnItemError ?? true,
            ComputeSourceHash = options?.ComputeSourceHash ?? false
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
            maxCachedBTreePages: source.MaxCachedBTreePages,
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
            maxMessageBytes: source.MaxMessageBytes,
            maxDirectoryDepth: source.MaxDirectoryDepth,
            maxDirectoryFileCount: source.MaxDirectoryFileCount);
    }

    private static EmailStoreItemReadOptions? CloneItemReadOptions(
        EmailStoreItemReadOptions? options) => options == null
        ? null
        : new EmailStoreItemReadOptions(options.Parts, options.MaxDecodedPropertyBytes,
            options.PreferStreamingAttachmentContent);

    private static EmailStoreQuery? CloneQuery(EmailStoreQuery? query) => query == null
        ? null
        : new EmailStoreQuery(
            query.FolderId,
            query.IncludeDescendants,
            query.IncludeAssociatedItems,
            query.IncludeOrphanedItems,
            query.ItemKind,
            query.SubjectContains,
            query.SenderContains,
            query.Since,
            query.Before,
            query.HasAttachments,
            query.IsRead,
            query.MaxItemsScanned,
            query.MaxResults);
}
