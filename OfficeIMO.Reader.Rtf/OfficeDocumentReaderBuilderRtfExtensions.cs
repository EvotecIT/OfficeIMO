namespace OfficeIMO.Reader.Rtf;

/// <summary>
/// Adds RTF support to <see cref="OfficeDocumentReaderBuilder"/>.
/// </summary>
public static class OfficeDocumentReaderBuilderRtfExtensions {
    /// <summary>
    /// Stable handler identifier for RTF adapter registration.
    /// </summary>
    public const string HandlerId = "officeimo.reader.rtf";

    /// <summary>
    /// Adds RTF ingestion to an isolated reader builder.
    /// </summary>
    public static OfficeDocumentReaderBuilder AddRtfHandler(
        this OfficeDocumentReaderBuilder builder,
        ReaderRtfOptions? rtfOptions = null,
        bool replaceExisting = true) {
        if (builder == null) throw new ArgumentNullException(nameof(builder));
        return builder.AddHandler(CreateRegistration(rtfOptions), replaceExisting);
    }

    private static ReaderHandlerRegistration CreateRegistration(ReaderRtfOptions? rtfOptions) {
        ReaderRtfOptions? registeredOptions = ReaderRtfOptionsCloner.CloneNullable(rtfOptions);
        return new ReaderHandlerRegistration {
            Origin = ReaderHandlerOrigin.OfficeIMO,
            Id = HandlerId,
            DisplayName = "RTF Reader Adapter",
            Description = "Modular RTF adapter using OfficeIMO.Rtf semantic read model.",
            Kind = ReaderInputKind.Rtf,
            Extensions = new[] { ".rtf" },
            ReadPath = (path, readerOptions, ct) => RtfReaderAdapter.Read(
                rtfPath: path,
                readerOptions: readerOptions,
                rtfOptions: ReaderRtfOptionsCloner.CloneNullable(registeredOptions),
                cancellationToken: ct),
            ReadStream = (stream, sourceName, readerOptions, ct) => RtfReaderAdapter.Read(
                rtfStream: stream,
                sourceName: sourceName,
                readerOptions: readerOptions,
                rtfOptions: ReaderRtfOptionsCloner.CloneNullable(registeredOptions),
                cancellationToken: ct),
            ReadDocumentPath = (path, readerOptions, ct) => RtfReaderAdapter.ReadDocument(
                rtfPath: path,
                readerOptions: readerOptions,
                rtfOptions: ReaderRtfOptionsCloner.CloneNullable(registeredOptions),
                cancellationToken: ct),
            ReadDocumentStream = (stream, sourceName, readerOptions, ct) => RtfReaderAdapter.ReadDocument(
                rtfStream: stream,
                sourceName: sourceName,
                readerOptions: readerOptions,
                rtfOptions: ReaderRtfOptionsCloner.CloneNullable(registeredOptions),
                cancellationToken: ct)
        };
    }
}
