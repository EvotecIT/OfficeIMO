using OfficeIMO.OpenDocument;

namespace OfficeIMO.Reader.OpenDocument;

/// <summary>Adds native OpenDocument ingestion to <see cref="OfficeDocumentReaderBuilder"/>.</summary>
public static class OfficeDocumentReaderBuilderOpenDocumentExtensions {
    /// <summary>Stable handler identifier for the packaged and flat OpenDocument adapter.</summary>
    public const string HandlerId = "officeimo.reader.opendocument";

    /// <summary>Adds native OpenDocument ingestion for packaged and flat text, spreadsheets, presentations, and graphics.</summary>
    public static OfficeDocumentReaderBuilder AddOpenDocumentHandler(
        this OfficeDocumentReaderBuilder builder,
        ReaderOpenDocumentOptions? options = null,
        bool replaceExisting = false) {
        ArgumentNullException.ThrowIfNull(builder);
        ReaderHandlerRegistration registration = CreateRegistration((options ?? new ReaderOpenDocumentOptions()).Clone());
        return builder.AddHandler(registration, replaceExisting);
    }
    private static ReaderHandlerRegistration CreateRegistration(ReaderOpenDocumentOptions formatOptions) {
        return new ReaderHandlerRegistration {
            Origin = ReaderHandlerOrigin.OfficeIMO,
            Id = HandlerId,
            DisplayName = "OpenDocument Reader Adapter",
            Description = "Native extraction from packaged and flat OpenDocument files.",
            Kind = ReaderInputKind.OpenDocument,
            Extensions = new[] { ".odt", ".ods", ".odp", ".odg", ".fodt", ".fods", ".fodp", ".fodg" },
            ReadPath = (path, options, cancellationToken) =>
                OpenDocumentReaderAdapter.Read(path, options, formatOptions, cancellationToken),
            ReadStream = (stream, sourceName, options, cancellationToken) =>
                OpenDocumentReaderAdapter.Read(stream, sourceName, options, formatOptions, cancellationToken),
            ReadDocumentPath = (path, options, cancellationToken) =>
                OpenDocumentReaderAdapter.ReadDocument(path, options, formatOptions, cancellationToken),
            ReadDocumentStream = (stream, sourceName, options, cancellationToken) =>
                OpenDocumentReaderAdapter.ReadDocument(stream, sourceName, options, formatOptions, cancellationToken),
            DeterministicOutput = true,
            SupportsTextAndMetadata = true
        };
    }
}
