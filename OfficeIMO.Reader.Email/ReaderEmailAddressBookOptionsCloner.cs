using OfficeIMO.Email.AddressBook;

namespace OfficeIMO.Reader.Email;

internal static class ReaderEmailAddressBookOptionsCloner {
    internal static ReaderEmailAddressBookOptions CloneOrDefault(ReaderEmailAddressBookOptions? options) =>
        new ReaderEmailAddressBookOptions {
            AddressBookOptions = options?.AddressBookOptions,
            AddressListId = options?.AddressListId,
            Query = options?.Query,
            MaxEntries = options?.MaxEntries ?? 10_000,
            ContinueOnEntryError = options?.ContinueOnEntryError ?? true,
            IncludeMembershipValues = options?.IncludeMembershipValues ?? false,
            MaxMultiValueItems = options?.MaxMultiValueItems ?? 100,
            ComputeSourceHash = options?.ComputeSourceHash ?? false
        };

    internal static OfflineAddressBookReaderOptions CreateEffective(
        ReaderEmailAddressBookOptions options, ReaderOptions readerOptions) {
        OfflineAddressBookReaderOptions source = options.AddressBookOptions ?? OfflineAddressBookReaderOptions.Default;
        long maxInput = readerOptions.MaxInputBytes.HasValue
            ? Math.Min(source.MaxInputBytes, readerOptions.MaxInputBytes.Value)
            : source.MaxInputBytes;
        return new OfflineAddressBookReaderOptions(
            maxInput,
            source.MaxDiscoveredFiles,
            source.MaxDirectoryDepth,
            source.MaxMetadataBytes,
            source.MaxPropertiesPerTable,
            source.MaxRecordBytes,
            source.MaxStringBytes,
            source.MaxBinaryBytes,
            source.MaxValuesPerProperty,
            source.MaxDeclaredEntries,
            source.String8CodePage,
            source.RetainRawPropertyBytes);
    }

    internal static OfflineAddressBookSearchQuery CreateEffectiveQuery(
        OfflineAddressBookSearchQuery query, ReaderEmailAddressBookOptions options) =>
        new OfflineAddressBookSearchQuery(
            query.Terms,
            query.Fields,
            query.MatchMode,
            options.AddressListId ?? query.AddressListId,
            query.ObjectType,
            query.MaxEntriesScanned,
            Math.Min(query.MaxResults, options.MaxEntries),
            query.MaxSearchableCharactersPerEntry,
            query.SnippetCharacters,
            query.ProgressInterval,
            query.ContinueOnEntryError,
            query.ResumeFrom);
}
