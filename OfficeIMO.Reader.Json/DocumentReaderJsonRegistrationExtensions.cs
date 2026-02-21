namespace OfficeIMO.Reader.Json;

/// <summary>
/// Registration helpers for plugging JSON support into <see cref="DocumentReader"/>.
/// </summary>
public static class DocumentReaderJsonRegistrationExtensions {
    /// <summary>
    /// Stable handler identifier for JSON adapter registration.
    /// </summary>
    public const string HandlerId = "officeimo.reader.json";

    /// <summary>
    /// Registers JSON ingestion into <see cref="DocumentReader"/>.
    /// </summary>
    /// <param name="jsonOptions">Default parser options used by this handler.</param>
    /// <param name="replaceExisting">
    /// Defaults to true because this extension is already handled by the built-in plain text path.
    /// </param>
    [ReaderHandlerRegistrar(HandlerId)]
    public static void RegisterJsonHandler(JsonReadOptions? jsonOptions = null, bool replaceExisting = true) {
        var registered = Clone(jsonOptions);

        DocumentReader.RegisterHandler(new ReaderHandlerRegistration {
            Id = HandlerId,
            DisplayName = "JSON Reader Adapter",
            Description = "Modular JSON AST parser with path/type/value chunk output.",
            Kind = ReaderInputKind.Text,
            Extensions = new[] { ".json" },
            ReadPath = (path, readerOptions, ct) => DocumentReaderJsonExtensions.ReadJson(
                path: path,
                readerOptions: readerOptions,
                jsonOptions: Clone(registered),
                cancellationToken: ct),
            ReadStream = (stream, sourceName, readerOptions, ct) => DocumentReaderJsonExtensions.ReadJson(
                stream: stream,
                sourceName: sourceName,
                readerOptions: readerOptions,
                jsonOptions: Clone(registered),
                cancellationToken: ct)
        }, replaceExisting);
    }

    /// <summary>
    /// Unregisters JSON ingestion handler from <see cref="DocumentReader"/>.
    /// </summary>
    public static bool UnregisterJsonHandler() {
        return DocumentReader.UnregisterHandler(HandlerId);
    }

    private static JsonReadOptions? Clone(JsonReadOptions? options) {
        if (options == null) return null;

        return new JsonReadOptions {
            ChunkRows = options.ChunkRows,
            MaxDepth = options.MaxDepth,
            IncludeMarkdown = options.IncludeMarkdown
        };
    }
}
