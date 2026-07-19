using OfficeIMO.OpenDocument;

namespace OfficeIMO.Reader.OpenDocument;

/// <summary>Adds native OpenDocument ingestion to <see cref="OfficeDocumentReaderBuilder"/>.</summary>
public static class OfficeDocumentReaderBuilderOpenDocumentExtensions {
    /// <summary>Stable handler identifier for the ODT, ODS, and ODP adapter.</summary>
    public const string HandlerId = "officeimo.reader.opendocument";

    /// <summary>Adds native OpenDocument ingestion for <c>.odt</c>, <c>.ods</c>, and <c>.odp</c>.</summary>
    public static OfficeDocumentReaderBuilder AddOpenDocumentHandler(
        this OfficeDocumentReaderBuilder builder,
        ReaderOpenDocumentOptions? options = null,
        bool replaceExisting = false) {
        if (builder == null) throw new ArgumentNullException(nameof(builder));
        ReaderHandlerRegistration registration = CreateRegistration((options ?? new ReaderOpenDocumentOptions()).Clone());
        return builder.AddHandler(registration, replaceExisting);
    }
    private static ReaderHandlerRegistration CreateRegistration(ReaderOpenDocumentOptions formatOptions) {
        return new ReaderHandlerRegistration {
            Origin = ReaderHandlerOrigin.OfficeIMO,
            Id = HandlerId,
            DisplayName = "OpenDocument Reader Adapter",
            Description = "Native dependency-free ODT, ODS, and ODP extraction.",
            Kind = ReaderInputKind.OpenDocument,
            Extensions = new[] { ".odt", ".ods", ".odp" },
            ReadPath = (path, options, cancellationToken) =>
                OpenDocumentReaderAdapter.Read(path, options, formatOptions, cancellationToken),
            ReadStream = (stream, sourceName, options, cancellationToken) =>
                OpenDocumentReaderAdapter.Read(stream, sourceName, options, formatOptions, cancellationToken),
            DeterministicOutput = true
        };
    }
}
