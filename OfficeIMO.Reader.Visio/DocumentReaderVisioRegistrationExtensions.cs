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
        var registeredOptions = ReaderVisioOptionsCloner.CloneNullable(visioOptions);

        DocumentReader.RegisterHandler(new ReaderHandlerRegistration {
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
                cancellationToken: ct)
        }, replaceExisting);
    }

    /// <summary>
    /// Unregisters Visio ingestion handler from <see cref="DocumentReader"/>.
    /// </summary>
    public static bool UnregisterVisioHandler() {
        return DocumentReader.UnregisterHandler(HandlerId);
    }
}
