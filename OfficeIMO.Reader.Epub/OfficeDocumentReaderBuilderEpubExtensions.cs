using OfficeIMO.Epub;

namespace OfficeIMO.Reader.Epub;

/// <summary>
/// Adds EPUB support to <see cref="OfficeDocumentReaderBuilder"/>.
/// </summary>
public static class OfficeDocumentReaderBuilderEpubExtensions {
    /// <summary>
    /// Stable handler identifier for EPUB adapter registration.
    /// </summary>
    public const string HandlerId = "officeimo.reader.epub";

    /// <summary>
    /// Adds EPUB ingestion to an isolated reader builder.
    /// </summary>
    public static OfficeDocumentReaderBuilder AddEpubHandler(
        this OfficeDocumentReaderBuilder builder,
        EpubReadOptions? epubOptions = null,
        bool replaceExisting = false) {
        if (builder == null) throw new ArgumentNullException(nameof(builder));
        ReaderHandlerRegistration registration = CreateRegistration(epubOptions);
        return builder.AddHandler(registration, replaceExisting);
    }

    private static ReaderHandlerRegistration CreateRegistration(EpubReadOptions? epubOptions) {
        EpubReadOptions? registeredOptions = Clone(epubOptions);
        return new ReaderHandlerRegistration {
            Origin = ReaderHandlerOrigin.OfficeIMO,
            Id = HandlerId,
            DisplayName = "EPUB Reader Adapter",
            Description = "Modular EPUB adapter that emits chapter-oriented Reader chunks.",
            Kind = ReaderInputKind.Epub,
            Extensions = new[] { ".epub" },
            ReadPath = (path, readerOptions, ct) => EpubReaderAdapter.Read(
                epubPath: path,
                readerOptions: readerOptions,
                epubOptions: Clone(registeredOptions),
                cancellationToken: ct),
            ReadStream = (stream, sourceName, readerOptions, ct) => EpubReaderAdapter.Read(
                epubStream: stream,
                sourceName: sourceName,
                readerOptions: readerOptions,
                epubOptions: Clone(registeredOptions),
                cancellationToken: ct),
            ReadDocumentPath = (path, readerOptions, ct) => EpubReaderAdapter.ReadDocument(
                epubPath: path,
                readerOptions: readerOptions,
                epubOptions: Clone(registeredOptions),
                cancellationToken: ct),
            ReadDocumentStream = (stream, sourceName, readerOptions, ct) => EpubReaderAdapter.ReadDocument(
                epubStream: stream,
                sourceName: sourceName,
                readerOptions: readerOptions,
                epubOptions: Clone(registeredOptions),
                cancellationToken: ct)
        };
    }

    private static EpubReadOptions? Clone(EpubReadOptions? options) {
        if (options == null) return null;
        return new EpubReadOptions {
            MaxPackageBytes = options.MaxPackageBytes,
            MaxArchiveEntries = options.MaxArchiveEntries,
            MaxTotalUncompressedBytes = options.MaxTotalUncompressedBytes,
            MaxPackageMetadataBytes = options.MaxPackageMetadataBytes,
            MaxMetadataItems = options.MaxMetadataItems,
            MaxNavigationItems = options.MaxNavigationItems,
            MaxNavigationDepth = options.MaxNavigationDepth,
            MaxChapters = options.MaxChapters,
            MaxChapterBytes = options.MaxChapterBytes,
            MaxTotalRawHtmlBytes = options.MaxTotalRawHtmlBytes,
            IncludeRawHtml = options.IncludeRawHtml,
            IncludeResourceData = options.IncludeResourceData,
            MaxResources = options.MaxResources,
            MaxResourceBytes = options.MaxResourceBytes,
            MaxTotalResourceBytes = options.MaxTotalResourceBytes,
            DeterministicOrder = options.DeterministicOrder,
            PreferSpineOrder = options.PreferSpineOrder,
            IncludeNonLinearSpineItems = options.IncludeNonLinearSpineItems,
            FallbackToHtmlScan = options.FallbackToHtmlScan
        };
    }
}
