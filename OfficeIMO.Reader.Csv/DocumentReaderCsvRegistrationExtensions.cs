namespace OfficeIMO.Reader.Csv;

/// <summary>
/// Registration helpers for plugging CSV/TSV support into <see cref="DocumentReader"/>.
/// </summary>
public static class DocumentReaderCsvRegistrationExtensions {
    /// <summary>
    /// Stable handler identifier for CSV/TSV adapter registration.
    /// </summary>
    public const string HandlerId = "officeimo.reader.csv";

    /// <summary>
    /// Registers CSV/TSV ingestion into <see cref="DocumentReader"/>.
    /// </summary>
    /// <param name="csvOptions">Default parser options used by this handler.</param>
    /// <param name="replaceExisting">
    /// Defaults to true because these extensions are already handled by the built-in plain text path.
    /// </param>
    [ReaderHandlerRegistrar(HandlerId)]
    public static void RegisterCsvHandler(CsvReadOptions? csvOptions = null, bool replaceExisting = true) {
        var registered = Clone(csvOptions);

        DocumentReader.RegisterHandler(new ReaderHandlerRegistration {
            Id = HandlerId,
            DisplayName = "CSV Reader Adapter",
            Description = "Modular CSV/TSV parser with table-aware chunk output.",
            Kind = ReaderInputKind.Csv,
            Extensions = new[] { ".csv", ".tsv" },
            ReadPath = (path, readerOptions, ct) => DocumentReaderCsvExtensions.ReadCsv(
                path: path,
                readerOptions: readerOptions,
                csvOptions: Clone(registered),
                cancellationToken: ct),
            ReadStream = (stream, sourceName, readerOptions, ct) => DocumentReaderCsvExtensions.ReadCsv(
                stream: stream,
                sourceName: sourceName,
                readerOptions: readerOptions,
                csvOptions: Clone(registered),
                cancellationToken: ct)
        }, replaceExisting);
    }

    /// <summary>
    /// Unregisters CSV/TSV ingestion handler from <see cref="DocumentReader"/>.
    /// </summary>
    public static bool UnregisterCsvHandler() {
        return DocumentReader.UnregisterHandler(HandlerId);
    }

    private static CsvReadOptions? Clone(CsvReadOptions? options) {
        if (options == null) return null;

        return new CsvReadOptions {
            ChunkRows = options.ChunkRows,
            HeadersInFirstRow = options.HeadersInFirstRow,
            IncludeMarkdown = options.IncludeMarkdown
        };
    }
}
