namespace OfficeIMO.Reader.Json;

/// <summary>
/// Adds JSON support to <see cref="OfficeDocumentReaderBuilder"/>.
/// </summary>
public static class OfficeDocumentReaderBuilderJsonExtensions {
    /// <summary>
    /// Stable handler identifier for JSON adapter registration.
    /// </summary>
    public const string HandlerId = "officeimo.reader.json";

    /// <summary>
    /// Adds JSON ingestion to an isolated reader builder.
    /// </summary>
    public static OfficeDocumentReaderBuilder AddJsonHandler(
        this OfficeDocumentReaderBuilder builder,
        JsonReadOptions? jsonOptions = null,
        bool replaceExisting = true) {
        if (builder == null) throw new ArgumentNullException(nameof(builder));
        ReaderHandlerRegistration registration = CreateRegistration(jsonOptions);
        return builder.AddHandler(registration, replaceExisting);
    }

    private static ReaderHandlerRegistration CreateRegistration(JsonReadOptions? jsonOptions) {
        JsonReadOptions? registered = Clone(jsonOptions);
        return new ReaderHandlerRegistration {
            Origin = ReaderHandlerOrigin.OfficeIMO,
            Id = HandlerId,
            DisplayName = "JSON Reader Adapter",
            Description = "Modular JSON AST parser with path/type/value chunk output.",
            Kind = ReaderInputKind.Json,
            Extensions = new[] { ".json" },
            ReadPath = (path, readerOptions, ct) => JsonReaderAdapter.Read(
                path: path,
                readerOptions: readerOptions,
                jsonOptions: Clone(registered),
                cancellationToken: ct),
            ReadStream = (stream, sourceName, readerOptions, ct) => JsonReaderAdapter.Read(
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
