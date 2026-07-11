namespace OfficeIMO.Reader.Html;

/// <summary>
/// Registration helpers for plugging HTML support into <see cref="DocumentReader"/>.
/// </summary>
public static class DocumentReaderHtmlRegistrationExtensions {
    /// <summary>
    /// Stable handler identifier for HTML adapter registration.
    /// </summary>
    public const string HandlerId = "officeimo.reader.html";

    /// <summary>
    /// Registers HTML ingestion into <see cref="DocumentReader"/> for <c>.html</c>, <c>.htm</c>, and <c>.xhtml</c>.
    /// </summary>
    [ReaderHandlerRegistrar(HandlerId)]
    public static void RegisterHtmlHandler(ReaderHtmlOptions? htmlOptions = null, bool replaceExisting = false) {
        RegisterHtmlHandler(htmlOptions, replaceExisting, preserveExistingCustomExtensions: false);
    }

    /// <summary>
    /// Registers HTML ingestion into <see cref="DocumentReader"/> for <c>.html</c>, <c>.htm</c>, and <c>.xhtml</c>.
    /// </summary>
    public static void RegisterHtmlHandler(ReaderHtmlOptions? htmlOptions, bool replaceExisting, bool preserveExistingCustomExtensions) {
        ReaderHandlerRegistration registration = CreateRegistration(htmlOptions);

        if (preserveExistingCustomExtensions) {
            DocumentReader.RegisterHandlerPreservingExistingCustomExtensions(registration, replaceExisting);
        } else {
            DocumentReader.RegisterHandler(registration, replaceExisting);
        }
    }

    /// <summary>
    /// Adds HTML ingestion to an isolated reader builder.
    /// </summary>
    public static OfficeDocumentReaderBuilder AddHtmlHandler(
        this OfficeDocumentReaderBuilder builder,
        ReaderHtmlOptions? htmlOptions = null,
        bool replaceExisting = false,
        bool preserveExistingCustomExtensions = false) {
        if (builder == null) throw new ArgumentNullException(nameof(builder));
        ReaderHandlerRegistration registration = CreateRegistration(htmlOptions);
        if (preserveExistingCustomExtensions) {
            builder.AddHandlerPreservingExistingCustomExtensions(registration, replaceExisting);
        } else {
            builder.AddHandler(registration, replaceExisting);
        }
        return builder;
    }

    /// <summary>
    /// Unregisters HTML ingestion handler from <see cref="DocumentReader"/>.
    /// </summary>
    public static bool UnregisterHtmlHandler() {
        return DocumentReader.UnregisterHandler(HandlerId);
    }

    private static ReaderHandlerRegistration CreateRegistration(ReaderHtmlOptions? htmlOptions) {
        ReaderHtmlOptions? registeredOptions = ReaderHtmlOptionsCloner.CloneNullable(htmlOptions);
        return new ReaderHandlerRegistration {
            Id = HandlerId,
            DisplayName = "HTML Reader Adapter",
            Description = "Modular HTML adapter using OfficeIMO.Markdown.Html.",
            Kind = ReaderInputKind.Html,
            Extensions = new[] { ".html", ".htm", ".xhtml" },
            ReadPath = (path, readerOptions, ct) => DocumentReaderHtmlExtensions.ReadHtmlFile(
                htmlPath: path,
                readerOptions: readerOptions,
                htmlOptions: ReaderHtmlOptionsCloner.CloneNullable(registeredOptions),
                cancellationToken: ct),
            ReadStream = (stream, sourceName, readerOptions, ct) => DocumentReaderHtmlExtensions.ReadHtml(
                htmlStream: stream,
                sourceName: sourceName,
                readerOptions: readerOptions,
                htmlOptions: ReaderHtmlOptionsCloner.CloneNullable(registeredOptions),
                cancellationToken: ct),
            ReadDocumentPath = (path, readerOptions, ct) => DocumentReaderHtmlExtensions.ReadHtmlDocument(
                htmlPath: path,
                readerOptions: readerOptions,
                htmlOptions: ReaderHtmlOptionsCloner.CloneNullable(registeredOptions),
                cancellationToken: ct),
            ReadDocumentStream = (stream, sourceName, readerOptions, ct) => DocumentReaderHtmlExtensions.ReadHtmlDocument(
                htmlStream: stream,
                sourceName: sourceName,
                readerOptions: readerOptions,
                htmlOptions: ReaderHtmlOptionsCloner.CloneNullable(registeredOptions),
                cancellationToken: ct)
        };
    }
}
