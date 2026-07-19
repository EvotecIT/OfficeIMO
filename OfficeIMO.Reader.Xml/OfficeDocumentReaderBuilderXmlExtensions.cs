namespace OfficeIMO.Reader.Xml;

/// <summary>
/// Adds XML support to <see cref="OfficeDocumentReaderBuilder"/>.
/// </summary>
public static class OfficeDocumentReaderBuilderXmlExtensions {
    /// <summary>
    /// Stable handler identifier for XML adapter registration.
    /// </summary>
    public const string HandlerId = "officeimo.reader.xml";

    /// <summary>
    /// Adds XML ingestion to an isolated reader builder.
    /// </summary>
    public static OfficeDocumentReaderBuilder AddXmlHandler(
        this OfficeDocumentReaderBuilder builder,
        XmlReadOptions? xmlOptions = null,
        bool replaceExisting = true) {
        if (builder == null) throw new ArgumentNullException(nameof(builder));
        ReaderHandlerRegistration registration = CreateRegistration(xmlOptions);
        return builder.AddHandler(registration, replaceExisting);
    }

    private static ReaderHandlerRegistration CreateRegistration(XmlReadOptions? xmlOptions) {
        XmlReadOptions? registered = Clone(xmlOptions);
        return new ReaderHandlerRegistration {
            Origin = ReaderHandlerOrigin.OfficeIMO,
            Id = HandlerId,
            DisplayName = "XML Reader Adapter",
            Description = "Modular XML tree parser with path/type/value chunk output.",
            Kind = ReaderInputKind.Xml,
            Extensions = new[] { ".xml" },
            ReadPath = (path, readerOptions, ct) => XmlReaderAdapter.Read(
                path: path,
                readerOptions: readerOptions,
                xmlOptions: Clone(registered),
                cancellationToken: ct),
            ReadStream = (stream, sourceName, readerOptions, ct) => XmlReaderAdapter.Read(
                stream: stream,
                sourceName: sourceName,
                readerOptions: readerOptions,
                xmlOptions: Clone(registered),
                cancellationToken: ct)
        };
    }

    private static XmlReadOptions? Clone(XmlReadOptions? options) {
        if (options == null) return null;

        return new XmlReadOptions {
            ChunkRows = options.ChunkRows,
            IncludeMarkdown = options.IncludeMarkdown
        };
    }
}
