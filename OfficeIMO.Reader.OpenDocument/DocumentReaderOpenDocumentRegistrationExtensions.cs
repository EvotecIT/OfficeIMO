using OfficeIMO.OpenDocument;

namespace OfficeIMO.Reader.OpenDocument;

/// <summary>Registration helpers for native OpenDocument ingestion.</summary>
public static class DocumentReaderOpenDocumentRegistrationExtensions {
    /// <summary>Stable handler identifier for the ODT, ODS, and ODP adapter.</summary>
    public const string HandlerId = "officeimo.reader.opendocument";

    /// <summary>Registers native OpenDocument ingestion for <c>.odt</c>, <c>.ods</c>, and <c>.odp</c>.</summary>
    [ReaderHandlerRegistrar(HandlerId)]
    public static void RegisterOpenDocumentHandler(bool replaceExisting = false) {
        var registration = new ReaderHandlerRegistration {
            Id = HandlerId,
            DisplayName = "OpenDocument Reader Adapter",
            Description = "Native dependency-free ODT, ODS, and ODP extraction.",
            Kind = ReaderInputKind.OpenDocument,
            Extensions = new[] { ".odt", ".ods", ".odp" },
            ReadPath = (path, options, cancellationToken) =>
                DocumentReaderOpenDocumentExtensions.ReadOpenDocument(path, options, cancellationToken),
            ReadStream = (stream, sourceName, options, cancellationToken) =>
                DocumentReaderOpenDocumentExtensions.ReadOpenDocument(stream, sourceName, options, cancellationToken),
            DeterministicOutput = true
        };
        DocumentReader.RegisterHandler(registration, replaceExisting);
    }

    /// <summary>Unregisters the native OpenDocument handler.</summary>
    public static bool UnregisterOpenDocumentHandler() => DocumentReader.UnregisterHandler(HandlerId);
}
