using OfficeIMO.Epub;

namespace OfficeIMO.Reader.Epub;

/// <summary>
/// Registration helpers for plugging EPUB support into <see cref="DocumentReader"/>.
/// </summary>
public static class DocumentReaderEpubRegistrationExtensions {
    /// <summary>
    /// Stable handler identifier for EPUB adapter registration.
    /// </summary>
    public const string HandlerId = "officeimo.reader.epub";

    /// <summary>
    /// Registers EPUB ingestion into <see cref="DocumentReader"/> for the <c>.epub</c> extension.
    /// </summary>
    [ReaderHandlerRegistrar(HandlerId)]
    public static void RegisterEpubHandler(EpubReadOptions? epubOptions = null, bool replaceExisting = false) {
        var registeredOptions = Clone(epubOptions);

        DocumentReader.RegisterHandler(new ReaderHandlerRegistration {
            Id = HandlerId,
            DisplayName = "EPUB Reader Adapter",
            Description = "Modular EPUB adapter that emits chapter-oriented Reader chunks.",
            Kind = ReaderInputKind.Unknown,
            Extensions = new[] { ".epub" },
            ReadPath = (path, readerOptions, ct) => DocumentReaderEpubExtensions.ReadEpub(
                epubPath: path,
                readerOptions: readerOptions,
                epubOptions: Clone(registeredOptions),
                cancellationToken: ct),
            ReadStream = (stream, sourceName, readerOptions, ct) => DocumentReaderEpubExtensions.ReadEpub(
                epubStream: stream,
                sourceName: sourceName,
                readerOptions: readerOptions,
                epubOptions: Clone(registeredOptions),
                cancellationToken: ct)
        }, replaceExisting);
    }

    /// <summary>
    /// Unregisters EPUB ingestion handler from <see cref="DocumentReader"/>.
    /// </summary>
    public static bool UnregisterEpubHandler() {
        return DocumentReader.UnregisterHandler(HandlerId);
    }

    private static EpubReadOptions? Clone(EpubReadOptions? options) {
        if (options == null) return null;
        return new EpubReadOptions {
            MaxChapters = options.MaxChapters,
            MaxChapterBytes = options.MaxChapterBytes,
            IncludeRawHtml = options.IncludeRawHtml,
            DeterministicOrder = options.DeterministicOrder,
            PreferSpineOrder = options.PreferSpineOrder,
            IncludeNonLinearSpineItems = options.IncludeNonLinearSpineItems,
            FallbackToHtmlScan = options.FallbackToHtmlScan
        };
    }
}
