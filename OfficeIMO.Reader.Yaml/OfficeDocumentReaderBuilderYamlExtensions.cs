namespace OfficeIMO.Reader.Yaml;

/// <summary>
/// Adds YAML support to <see cref="OfficeDocumentReaderBuilder"/>.
/// </summary>
public static class OfficeDocumentReaderBuilderYamlExtensions {
    /// <summary>
    /// Stable handler identifier for YAML adapter registration.
    /// </summary>
    public const string HandlerId = "officeimo.reader.yaml";

    /// <summary>
    /// Adds YAML ingestion to an isolated reader builder.
    /// </summary>
    public static OfficeDocumentReaderBuilder AddYamlHandler(
        this OfficeDocumentReaderBuilder builder,
        YamlReadOptions? yamlOptions = null,
        bool replaceExisting = true) {
        if (builder == null) throw new ArgumentNullException(nameof(builder));
        ReaderHandlerRegistration registration = CreateRegistration(yamlOptions);
        return builder.AddHandler(registration, replaceExisting);
    }

    private static ReaderHandlerRegistration CreateRegistration(YamlReadOptions? yamlOptions) {
        YamlReadOptions? registered = Clone(yamlOptions);
        return new ReaderHandlerRegistration {
            Origin = ReaderHandlerOrigin.OfficeIMO,
            Id = HandlerId,
            DisplayName = "YAML Reader Adapter",
            Description = "Modular YAML parser with path/type/value chunk output.",
            Kind = ReaderInputKind.Yaml,
            Extensions = new[] { ".yaml", ".yml" },
            ReadPath = (path, readerOptions, ct) => YamlReaderAdapter.Read(
                path: path,
                readerOptions: readerOptions,
                yamlOptions: Clone(registered),
                cancellationToken: ct),
            ReadStream = (stream, sourceName, readerOptions, ct) => YamlReaderAdapter.Read(
                stream: stream,
                sourceName: sourceName,
                readerOptions: readerOptions,
                yamlOptions: Clone(registered),
                cancellationToken: ct)
        };
    }

    private static YamlReadOptions? Clone(YamlReadOptions? options) {
        if (options == null) return null;

        return new YamlReadOptions {
            ChunkRows = options.ChunkRows,
            MaxDepth = options.MaxDepth,
            MaxNodes = options.MaxNodes,
            MaxParseEvents = options.MaxParseEvents,
            MaxScalarLength = options.MaxScalarLength,
            IncludeMarkdown = options.IncludeMarkdown
        };
    }
}
