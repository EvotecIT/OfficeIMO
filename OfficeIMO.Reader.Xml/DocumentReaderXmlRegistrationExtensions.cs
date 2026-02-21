namespace OfficeIMO.Reader.Xml;

/// <summary>
/// Registration helpers for plugging XML support into <see cref="DocumentReader"/>.
/// </summary>
public static class DocumentReaderXmlRegistrationExtensions {
    /// <summary>
    /// Stable handler identifier for XML adapter registration.
    /// </summary>
    public const string HandlerId = "officeimo.reader.xml";

    /// <summary>
    /// Registers XML ingestion into <see cref="DocumentReader"/>.
    /// </summary>
    /// <param name="xmlOptions">Default parser options used by this handler.</param>
    /// <param name="replaceExisting">
    /// Defaults to true because this extension is already handled by the built-in plain text path.
    /// </param>
    [ReaderHandlerRegistrar(HandlerId)]
    public static void RegisterXmlHandler(XmlReadOptions? xmlOptions = null, bool replaceExisting = true) {
        var registered = Clone(xmlOptions);

        DocumentReader.RegisterHandler(new ReaderHandlerRegistration {
            Id = HandlerId,
            DisplayName = "XML Reader Adapter",
            Description = "Modular XML tree parser with path/type/value chunk output.",
            Kind = ReaderInputKind.Text,
            Extensions = new[] { ".xml" },
            ReadPath = (path, readerOptions, ct) => DocumentReaderXmlExtensions.ReadXml(
                path: path,
                readerOptions: readerOptions,
                xmlOptions: Clone(registered),
                cancellationToken: ct),
            ReadStream = (stream, sourceName, readerOptions, ct) => DocumentReaderXmlExtensions.ReadXml(
                stream: stream,
                sourceName: sourceName,
                readerOptions: readerOptions,
                xmlOptions: Clone(registered),
                cancellationToken: ct)
        }, replaceExisting);
    }

    /// <summary>
    /// Unregisters XML ingestion handler from <see cref="DocumentReader"/>.
    /// </summary>
    public static bool UnregisterXmlHandler() {
        return DocumentReader.UnregisterHandler(HandlerId);
    }

    private static XmlReadOptions? Clone(XmlReadOptions? options) {
        if (options == null) return null;

        return new XmlReadOptions {
            ChunkRows = options.ChunkRows,
            IncludeMarkdown = options.IncludeMarkdown
        };
    }
}
