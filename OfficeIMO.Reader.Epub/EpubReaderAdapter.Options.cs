using OfficeIMO.Epub;

namespace OfficeIMO.Reader.Epub;

internal static partial class EpubReaderAdapter {
    /// <summary>
    /// Creates Reader-owned EPUB options without mutating caller-owned configuration.
    /// Reader always needs chapter XHTML internally to preserve document structure.
    /// </summary>
    private static EpubReadOptions CreateStructuredOptions(EpubReadOptions? options) {
        EpubReadOptions source = options ?? new EpubReadOptions();
        return CloneOptions(
            source,
            includeRawHtml: true,
            includeResourceData: source.IncludeResourceData);
    }

    /// <summary>
    /// Creates options for the rich result, which includes bounded manifest payloads by default.
    /// </summary>
    private static EpubReadOptions CreateRichOptions(EpubReadOptions? options) {
        EpubReadOptions source = options ?? new EpubReadOptions();
        return CloneOptions(
            source,
            includeRawHtml: true,
            includeResourceData: options == null || source.IncludeResourceData);
    }

    private static EpubReadOptions CloneOptions(
        EpubReadOptions source,
        bool includeRawHtml,
        bool includeResourceData) {
        return new EpubReadOptions {
            MaxPackageBytes = source.MaxPackageBytes,
            MaxArchiveEntries = source.MaxArchiveEntries,
            MaxTotalUncompressedBytes = source.MaxTotalUncompressedBytes,
            MaxPackageMetadataBytes = source.MaxPackageMetadataBytes,
            MaxMetadataItems = source.MaxMetadataItems,
            MaxNavigationItems = source.MaxNavigationItems,
            MaxNavigationDepth = source.MaxNavigationDepth,
            MaxChapters = source.MaxChapters,
            MaxChapterBytes = source.MaxChapterBytes,
            MaxTotalRawHtmlBytes = source.MaxTotalRawHtmlBytes,
            IncludeRawHtml = includeRawHtml,
            IncludeResourceData = includeResourceData,
            MaxResources = source.MaxResources,
            MaxResourceBytes = source.MaxResourceBytes,
            MaxTotalResourceBytes = source.MaxTotalResourceBytes,
            DeterministicOrder = source.DeterministicOrder,
            PreferSpineOrder = source.PreferSpineOrder,
            IncludeNonLinearSpineItems = source.IncludeNonLinearSpineItems,
            FallbackToHtmlScan = source.FallbackToHtmlScan
        };
    }
}
