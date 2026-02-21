namespace OfficeIMO.Reader.Text;

/// <summary>
/// Registration helpers for plugging structured text support into <see cref="DocumentReader"/>.
/// </summary>
public static class DocumentReaderTextRegistrationExtensions {
    /// <summary>
    /// Stable handler identifier for structured text adapter registration.
    /// </summary>
    public const string HandlerId = "officeimo.reader.text.structured";

    /// <summary>
    /// Registers structured CSV/JSON/XML ingestion into <see cref="DocumentReader"/>.
    /// </summary>
    /// <param name="structuredOptions">Default structured parser options used by this handler.</param>
    /// <param name="replaceExisting">
    /// Defaults to true because these extensions are already handled by the built-in plain text path.
    /// </param>
    public static void RegisterStructuredTextHandler(StructuredTextReadOptions? structuredOptions = null, bool replaceExisting = true) {
        var registered = Clone(structuredOptions);

        DocumentReader.RegisterHandler(new ReaderHandlerRegistration {
            Id = HandlerId,
            DisplayName = "Structured Text Reader Adapter",
            Description = "Modular structured parser for CSV/JSON/XML with table-aware chunk output.",
            Kind = ReaderInputKind.Text,
            Extensions = new[] { ".csv", ".tsv", ".json", ".xml" },
            ReadPath = (path, readerOptions, ct) => DocumentReaderTextExtensions.ReadStructuredText(
                path: path,
                readerOptions: readerOptions,
                structuredOptions: Clone(registered),
                cancellationToken: ct)
        }, replaceExisting);
    }

    /// <summary>
    /// Unregisters structured text ingestion handler from <see cref="DocumentReader"/>.
    /// </summary>
    public static bool UnregisterStructuredTextHandler() {
        return DocumentReader.UnregisterHandler(HandlerId);
    }

    private static StructuredTextReadOptions? Clone(StructuredTextReadOptions? options) {
        if (options == null) return null;
        return new StructuredTextReadOptions {
            CsvChunkRows = options.CsvChunkRows,
            CsvHeadersInFirstRow = options.CsvHeadersInFirstRow,
            IncludeCsvMarkdown = options.IncludeCsvMarkdown,
            JsonChunkRows = options.JsonChunkRows,
            JsonMaxDepth = options.JsonMaxDepth,
            IncludeJsonMarkdown = options.IncludeJsonMarkdown,
            XmlChunkRows = options.XmlChunkRows,
            IncludeXmlMarkdown = options.IncludeXmlMarkdown
        };
    }
}
