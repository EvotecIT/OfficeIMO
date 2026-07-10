namespace OfficeIMO.Reader.Yaml;

/// <summary>
/// Registration helpers for plugging YAML support into <see cref="DocumentReader"/>.
/// </summary>
public static class DocumentReaderYamlRegistrationExtensions {
    /// <summary>
    /// Stable handler identifier for YAML adapter registration.
    /// </summary>
    public const string HandlerId = "officeimo.reader.yaml";

    /// <summary>
    /// Registers YAML ingestion into <see cref="DocumentReader"/>.
    /// </summary>
    /// <param name="yamlOptions">Default parser options used by this handler.</param>
    /// <param name="replaceExisting">
    /// Defaults to true because these extensions are already handled by the built-in plain text path.
    /// </param>
    /// <param name="preserveExistingCustomExtensions">When true, leaves extensions already owned by other custom handlers untouched.</param>
    [ReaderHandlerRegistrar(HandlerId)]
    public static void RegisterYamlHandler(YamlReadOptions? yamlOptions = null, bool replaceExisting = true, bool preserveExistingCustomExtensions = false) {
        ReaderHandlerRegistration registration = CreateRegistration(yamlOptions);

        if (preserveExistingCustomExtensions) {
            DocumentReader.RegisterHandlerPreservingExistingCustomExtensions(registration, replaceExisting);
        } else {
            DocumentReader.RegisterHandler(registration, replaceExisting);
        }
    }

    /// <summary>
    /// Adds YAML ingestion to an isolated reader builder.
    /// </summary>
    public static OfficeDocumentReaderBuilder AddYamlHandler(
        this OfficeDocumentReaderBuilder builder,
        YamlReadOptions? yamlOptions = null,
        bool replaceExisting = true,
        bool preserveExistingCustomExtensions = false) {
        if (builder == null) throw new ArgumentNullException(nameof(builder));
        ReaderHandlerRegistration registration = CreateRegistration(yamlOptions);
        if (preserveExistingCustomExtensions) {
            builder.AddHandlerPreservingExistingCustomExtensions(registration, replaceExisting);
        } else {
            builder.AddHandler(registration, replaceExisting);
        }
        return builder;
    }

    /// <summary>
    /// Unregisters YAML ingestion handler from <see cref="DocumentReader"/>.
    /// </summary>
    public static bool UnregisterYamlHandler() {
        return DocumentReader.UnregisterHandler(HandlerId);
    }

    private static ReaderHandlerRegistration CreateRegistration(YamlReadOptions? yamlOptions) {
        YamlReadOptions? registered = Clone(yamlOptions);
        return new ReaderHandlerRegistration {
            Id = HandlerId,
            DisplayName = "YAML Reader Adapter",
            Description = "Modular YAML parser with path/type/value chunk output.",
            Kind = ReaderInputKind.Yaml,
            Extensions = new[] { ".yaml", ".yml" },
            ReadPath = (path, readerOptions, ct) => DocumentReaderYamlExtensions.ReadYaml(
                path: path,
                readerOptions: readerOptions,
                yamlOptions: Clone(registered),
                cancellationToken: ct),
            ReadStream = (stream, sourceName, readerOptions, ct) => DocumentReaderYamlExtensions.ReadYaml(
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
