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
        RegisterJsonHandler(jsonOptions, replaceExisting, preserveExistingCustomExtensions: false);
    }

    /// <summary>
    /// Registers JSON ingestion into <see cref="DocumentReader"/>.
    /// </summary>
    /// <param name="jsonOptions">Default parser options used by this handler.</param>
    /// <param name="replaceExisting">
    /// Defaults to true because this extension is already handled by the built-in plain text path.
    /// </param>
    /// <param name="preserveExistingCustomExtensions">When true, leaves extensions already owned by other custom handlers untouched.</param>
    public static void RegisterJsonHandler(JsonReadOptions? jsonOptions, bool replaceExisting, bool preserveExistingCustomExtensions) {
        ReaderHandlerRegistration registration = CreateRegistration(jsonOptions);

        if (preserveExistingCustomExtensions) {
            DocumentReader.RegisterHandlerPreservingExistingCustomExtensions(registration, replaceExisting);
        } else {
            DocumentReader.RegisterHandler(registration, replaceExisting);
        }
    }

    /// <summary>
    /// Adds JSON ingestion to an isolated reader builder.
    /// </summary>
    public static OfficeDocumentReaderBuilder AddJsonHandler(
        this OfficeDocumentReaderBuilder builder,
        JsonReadOptions? jsonOptions = null,
        bool replaceExisting = true,
        bool preserveExistingCustomExtensions = false) {
        if (builder == null) throw new ArgumentNullException(nameof(builder));
        ReaderHandlerRegistration registration = CreateRegistration(jsonOptions);
        if (preserveExistingCustomExtensions) {
            builder.AddHandlerPreservingExistingCustomExtensions(registration, replaceExisting);
        } else {
            builder.AddHandler(registration, replaceExisting);
        }
        return builder;
    }

    /// <summary>
    /// Unregisters JSON ingestion handler from <see cref="DocumentReader"/>.
    /// </summary>
    public static bool UnregisterJsonHandler() {
        return DocumentReader.UnregisterHandler(HandlerId);
    }

    private static ReaderHandlerRegistration CreateRegistration(JsonReadOptions? jsonOptions) {
        JsonReadOptions? registered = Clone(jsonOptions);
        return new ReaderHandlerRegistration {
            Id = HandlerId,
            DisplayName = "JSON Reader Adapter",
            Description = "Modular JSON AST parser with path/type/value chunk output.",
            Kind = ReaderInputKind.Json,
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
        };
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
