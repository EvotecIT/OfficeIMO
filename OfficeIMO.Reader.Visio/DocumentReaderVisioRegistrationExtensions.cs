namespace OfficeIMO.Reader.Visio;

/// <summary>
/// Registration helpers for plugging Visio support into <see cref="DocumentReader"/>.
/// </summary>
public static class DocumentReaderVisioRegistrationExtensions {
    /// <summary>
    /// Stable handler identifier for Visio adapter registration.
    /// </summary>
    public const string HandlerId = "officeimo.reader.visio";

    /// <summary>
    /// Registers Visio ingestion into <see cref="DocumentReader"/> for VSDX/VSDM/VSTX/VSTM files and streams.
    /// </summary>
    [ReaderHandlerRegistrar(HandlerId)]
    public static void RegisterVisioHandler(ReaderVisioOptions? visioOptions = null, bool replaceExisting = false) {
        RegisterVisioHandler(visioOptions, replaceExisting, preserveExistingCustomExtensions: false);
    }

    /// <summary>
    /// Registers Visio ingestion into <see cref="DocumentReader"/> for VSDX/VSDM/VSTX/VSTM files and streams.
    /// </summary>
    public static void RegisterVisioHandler(ReaderVisioOptions? visioOptions, bool replaceExisting, bool preserveExistingCustomExtensions) {
        ReaderHandlerRegistration registration = CreateRegistration(visioOptions);

        if (preserveExistingCustomExtensions) {
            DocumentReader.RegisterHandlerPreservingExistingCustomExtensions(registration, replaceExisting);
        } else {
            DocumentReader.RegisterHandler(registration, replaceExisting);
        }
    }

    /// <summary>
    /// Adds Visio ingestion to an isolated reader builder.
    /// </summary>
    public static OfficeDocumentReaderBuilder AddVisioHandler(
        this OfficeDocumentReaderBuilder builder,
        ReaderVisioOptions? visioOptions = null,
        bool replaceExisting = false,
        bool preserveExistingCustomExtensions = false) {
        if (builder == null) throw new ArgumentNullException(nameof(builder));
        ReaderHandlerRegistration registration = CreateRegistration(visioOptions);
        if (preserveExistingCustomExtensions) {
            builder.AddHandlerPreservingExistingCustomExtensions(registration, replaceExisting);
        } else {
            builder.AddHandler(registration, replaceExisting);
        }
        return builder;
    }

    /// <summary>
    /// Unregisters Visio ingestion handler from <see cref="DocumentReader"/>.
    /// </summary>
    public static bool UnregisterVisioHandler() {
        return DocumentReader.UnregisterHandler(HandlerId);
    }

    private static ReaderHandlerRegistration CreateRegistration(ReaderVisioOptions? visioOptions) {
        ReaderVisioOptions? registeredOptions = ReaderVisioOptionsCloner.CloneNullable(visioOptions);
        return new ReaderHandlerRegistration {
            Id = HandlerId,
            DisplayName = "Visio Reader Adapter",
            Description = "Modular Visio adapter using OfficeIMO.Visio inspection snapshots.",
            Kind = ReaderInputKind.Visio,
            Extensions = new[] { ".vsdx", ".vsdm", ".vstx", ".vstm" },
            ReadPath = (path, readerOptions, ct) => DocumentReaderVisioExtensions.ReadVisioFile(
                visioPath: path,
                readerOptions: readerOptions,
                visioOptions: ReaderVisioOptionsCloner.CloneNullable(registeredOptions),
                cancellationToken: ct),
            ReadStream = (stream, sourceName, readerOptions, ct) => DocumentReaderVisioExtensions.ReadVisio(
                visioStream: stream,
                sourceName: sourceName,
                readerOptions: readerOptions,
                visioOptions: ReaderVisioOptionsCloner.CloneNullable(registeredOptions),
                cancellationToken: ct),
            ReadDocumentPath = (path, readerOptions, ct) => DocumentReaderVisioExtensions.ReadVisioDocument(
                visioPath: path,
                readerOptions: readerOptions,
                visioOptions: ReaderVisioOptionsCloner.CloneNullable(registeredOptions),
                cancellationToken: ct),
            ReadDocumentStream = (stream, sourceName, readerOptions, ct) => DocumentReaderVisioExtensions.ReadVisioDocument(
                visioStream: stream,
                sourceName: sourceName,
                readerOptions: readerOptions,
                visioOptions: ReaderVisioOptionsCloner.CloneNullable(registeredOptions),
                cancellationToken: ct)
        };
    }
}
