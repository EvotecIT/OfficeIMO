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
        RegisterEpubHandler(epubOptions, replaceExisting, preserveExistingCustomExtensions: false);
    }

    /// <summary>
    /// Registers EPUB ingestion into <see cref="DocumentReader"/> for the <c>.epub</c> extension.
    /// </summary>
    public static void RegisterEpubHandler(EpubReadOptions? epubOptions, bool replaceExisting, bool preserveExistingCustomExtensions) {
        ReaderHandlerRegistration registration = CreateRegistration(epubOptions);

        if (preserveExistingCustomExtensions) {
            DocumentReader.RegisterHandlerPreservingExistingCustomExtensions(registration, replaceExisting);
        } else {
            DocumentReader.RegisterHandler(registration, replaceExisting);
        }
    }

    /// <summary>
    /// Adds EPUB ingestion to an isolated reader builder.
    /// </summary>
    public static OfficeDocumentReaderBuilder AddEpubHandler(
        this OfficeDocumentReaderBuilder builder,
        EpubReadOptions? epubOptions = null,
        bool replaceExisting = false,
        bool preserveExistingCustomExtensions = false) {
        if (builder == null) throw new ArgumentNullException(nameof(builder));
        ReaderHandlerRegistration registration = CreateRegistration(epubOptions);
        if (preserveExistingCustomExtensions) {
            builder.AddHandlerPreservingExistingCustomExtensions(registration, replaceExisting);
        } else {
            builder.AddHandler(registration, replaceExisting);
        }
        return builder;
    }

    /// <summary>
    /// Unregisters EPUB ingestion handler from <see cref="DocumentReader"/>.
    /// </summary>
    public static bool UnregisterEpubHandler() {
        return DocumentReader.UnregisterHandler(HandlerId);
    }

    private static ReaderHandlerRegistration CreateRegistration(EpubReadOptions? epubOptions) {
        EpubReadOptions? registeredOptions = Clone(epubOptions);
        return new ReaderHandlerRegistration {
            Id = HandlerId,
            DisplayName = "EPUB Reader Adapter",
            Description = "Modular EPUB adapter that emits chapter-oriented Reader chunks.",
            Kind = ReaderInputKind.Epub,
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
                cancellationToken: ct),
            ReadDocumentPath = (path, readerOptions, ct) => DocumentReaderEpubExtensions.ReadEpubDocument(
                epubPath: path,
                readerOptions: readerOptions,
                epubOptions: Clone(registeredOptions),
                cancellationToken: ct),
            ReadDocumentStream = (stream, sourceName, readerOptions, ct) => DocumentReaderEpubExtensions.ReadEpubDocument(
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
            MaxChapters = options.MaxChapters,
            MaxChapterBytes = options.MaxChapterBytes,
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
