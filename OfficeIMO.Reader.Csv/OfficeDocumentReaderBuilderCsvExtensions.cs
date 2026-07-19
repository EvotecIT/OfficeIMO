namespace OfficeIMO.Reader.Csv;

/// <summary>
/// Adds CSV/TSV support to <see cref="OfficeDocumentReaderBuilder"/>.
/// </summary>
public static class OfficeDocumentReaderBuilderCsvExtensions {
    /// <summary>
    /// Stable handler identifier for CSV/TSV adapter registration.
    /// </summary>
    public const string HandlerId = "officeimo.reader.csv";

    /// <summary>
    /// Adds CSV/TSV ingestion to an isolated reader builder.
    /// </summary>
    public static OfficeDocumentReaderBuilder AddCsvHandler(
        this OfficeDocumentReaderBuilder builder,
        CsvReadOptions? csvOptions = null,
        bool replaceExisting = true) {
        if (builder == null) throw new ArgumentNullException(nameof(builder));
        ReaderHandlerRegistration registration = CreateRegistration(csvOptions);
        return builder.AddHandler(registration, replaceExisting);
    }

    private static ReaderHandlerRegistration CreateRegistration(CsvReadOptions? csvOptions) {
        CsvReadOptions? registered = Clone(csvOptions);
        return new ReaderHandlerRegistration {
            Origin = ReaderHandlerOrigin.OfficeIMO,
            Id = HandlerId,
            DisplayName = "CSV Reader Adapter",
            Description = "Modular CSV/TSV parser with table-aware chunk output.",
            Kind = ReaderInputKind.Csv,
            Extensions = new[] { ".csv", ".tsv" },
            ReadPath = (path, readerOptions, ct) => CsvReaderAdapter.Read(
                path: path,
                readerOptions: readerOptions,
                csvOptions: Clone(registered),
                cancellationToken: ct),
            ReadStream = (stream, sourceName, readerOptions, ct) => CsvReaderAdapter.Read(
                stream: stream,
                sourceName: sourceName,
                readerOptions: readerOptions,
                csvOptions: Clone(registered),
                cancellationToken: ct)
        };
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
