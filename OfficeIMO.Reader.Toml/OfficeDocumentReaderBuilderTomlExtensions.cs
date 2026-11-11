namespace OfficeIMO.Reader.Toml;

/// <summary>
/// Adds TOML support to <see cref="OfficeDocumentReaderBuilder"/>.
/// </summary>
public static class OfficeDocumentReaderBuilderTomlExtensions {
    /// <summary>
    /// Stable handler identifier for TOML adapter registration.
    /// </summary>
    public const string HandlerId = "officeimo.reader.toml";

    /// <summary>
    /// Adds TOML ingestion to an isolated reader builder.
    /// </summary>
    public static OfficeDocumentReaderBuilder AddTomlHandler(
        this OfficeDocumentReaderBuilder builder,
        TomlReadOptions? tomlOptions = null,
        bool replaceExisting = true) {
        ArgumentNullException.ThrowIfNull(builder);
        TomlReadOptions? registered = Clone(tomlOptions);
        return builder.AddHandler(new ReaderHandlerRegistration {
            Origin = ReaderHandlerOrigin.OfficeIMO,
            Id = HandlerId,
            DisplayName = "TOML Reader Adapter",
            Description = "TOML 1.1 configuration parser with path/type/value chunk output.",
            Kind = ReaderInputKind.Toml,
            Extensions = new[] { ".toml" },
            ReadPath = (path, readerOptions, ct) => TomlReaderAdapter.Read(path, readerOptions, Clone(registered), ct),
            ReadStream = (stream, sourceName, readerOptions, ct) => TomlReaderAdapter.Read(stream, sourceName, readerOptions, Clone(registered), ct)
        }, replaceExisting);
    }

    private static TomlReadOptions? Clone(TomlReadOptions? options) {
        if (options == null) return null;
        return new TomlReadOptions {
            ChunkRows = options.ChunkRows,
            IncludeMarkdown = options.IncludeMarkdown,
            MaxDepth = options.MaxDepth,
            MaxNodes = options.MaxNodes,
            MaxScalarLength = options.MaxScalarLength,
            LoadOptions = TomlReaderAdapter.CopyLoadOptions(options.LoadOptions)
        };
    }
}
