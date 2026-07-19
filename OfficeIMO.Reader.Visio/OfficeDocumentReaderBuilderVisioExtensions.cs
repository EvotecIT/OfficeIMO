namespace OfficeIMO.Reader.Visio;

/// <summary>
/// Adds Visio support to <see cref="OfficeDocumentReaderBuilder"/>.
/// </summary>
public static class OfficeDocumentReaderBuilderVisioExtensions {
    /// <summary>
    /// Stable handler identifier for Visio adapter registration.
    /// </summary>
    public const string HandlerId = "officeimo.reader.visio";

    /// <summary>
    /// Adds Visio ingestion to an isolated reader builder.
    /// </summary>
    public static OfficeDocumentReaderBuilder AddVisioHandler(
        this OfficeDocumentReaderBuilder builder,
        ReaderVisioOptions? visioOptions = null,
        bool replaceExisting = false) {
        if (builder == null) throw new ArgumentNullException(nameof(builder));
        ReaderHandlerRegistration registration = CreateRegistration(visioOptions);
        return builder.AddHandler(registration, replaceExisting);
    }

    private static ReaderHandlerRegistration CreateRegistration(ReaderVisioOptions? visioOptions) {
        ReaderVisioOptions? registeredOptions = ReaderVisioOptionsCloner.CloneNullable(visioOptions);
        return new ReaderHandlerRegistration {
            Origin = ReaderHandlerOrigin.OfficeIMO,
            Id = HandlerId,
            DisplayName = "Visio Reader Adapter",
            Description = "Modular Visio adapter using OfficeIMO.Visio inspection snapshots.",
            Kind = ReaderInputKind.Visio,
            Extensions = new[] { ".vsdx", ".vsdm", ".vstx", ".vstm" },
            ReadPath = (path, readerOptions, ct) => VisioReaderAdapter.Read(
                visioPath: path,
                readerOptions: readerOptions,
                visioOptions: ReaderVisioOptionsCloner.CloneNullable(registeredOptions),
                cancellationToken: ct),
            ReadStream = (stream, sourceName, readerOptions, ct) => VisioReaderAdapter.Read(
                visioStream: stream,
                sourceName: sourceName,
                readerOptions: readerOptions,
                visioOptions: ReaderVisioOptionsCloner.CloneNullable(registeredOptions),
                cancellationToken: ct),
            ReadDocumentPath = (path, readerOptions, ct) => VisioReaderAdapter.ReadDocument(
                visioPath: path,
                readerOptions: readerOptions,
                visioOptions: ReaderVisioOptionsCloner.CloneNullable(registeredOptions),
                cancellationToken: ct),
            ReadDocumentStream = (stream, sourceName, readerOptions, ct) => VisioReaderAdapter.ReadDocument(
                visioStream: stream,
                sourceName: sourceName,
                readerOptions: readerOptions,
                visioOptions: ReaderVisioOptionsCloner.CloneNullable(registeredOptions),
                cancellationToken: ct)
        };
    }
}
