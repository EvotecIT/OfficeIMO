namespace OfficeIMO.Reader.Pdf;

/// <summary>
/// Registration helpers for plugging PDF support into <see cref="DocumentReader"/>.
/// </summary>
public static class DocumentReaderPdfRegistrationExtensions {
    /// <summary>
    /// Stable handler identifier for PDF adapter registration.
    /// </summary>
    public const string HandlerId = "officeimo.reader.pdf";

    /// <summary>
    /// Registers PDF ingestion into <see cref="DocumentReader"/> for <c>.pdf</c> files and streams.
    /// </summary>
    [ReaderHandlerRegistrar(HandlerId)]
    public static void RegisterPdfHandler(ReaderPdfOptions? pdfOptions = null, bool replaceExisting = true) {
        var registeredOptions = ReaderPdfOptionsCloner.CloneNullable(pdfOptions);

        DocumentReader.RegisterHandler(new ReaderHandlerRegistration {
            Id = HandlerId,
            DisplayName = "PDF Reader Adapter",
            Description = "Modular PDF adapter using OfficeIMO.Pdf logical read model.",
            Kind = ReaderInputKind.Pdf,
            Extensions = new[] { ".pdf" },
            ReadPath = (path, readerOptions, ct) => DocumentReaderPdfExtensions.ReadPdfFile(
                pdfPath: path,
                readerOptions: readerOptions,
                pdfOptions: ReaderPdfOptionsCloner.CloneNullable(registeredOptions),
                cancellationToken: ct),
            ReadStream = (stream, sourceName, readerOptions, ct) => DocumentReaderPdfExtensions.ReadPdf(
                pdfStream: stream,
                sourceName: sourceName,
                readerOptions: readerOptions,
                pdfOptions: ReaderPdfOptionsCloner.CloneNullable(registeredOptions),
                cancellationToken: ct)
        }, replaceExisting);
    }

    /// <summary>
    /// Unregisters PDF ingestion handler from <see cref="DocumentReader"/>.
    /// </summary>
    public static bool UnregisterPdfHandler() {
        return DocumentReader.UnregisterHandler(HandlerId);
    }
}
