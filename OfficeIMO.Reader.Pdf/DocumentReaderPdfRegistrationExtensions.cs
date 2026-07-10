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
        DocumentReader.RegisterHandler(CreateRegistration(pdfOptions), replaceExisting);
    }

    /// <summary>
    /// Adds PDF ingestion to an isolated reader builder.
    /// </summary>
    public static OfficeDocumentReaderBuilder AddPdfHandler(
        this OfficeDocumentReaderBuilder builder,
        ReaderPdfOptions? pdfOptions = null,
        bool replaceExisting = true) {
        if (builder == null) throw new ArgumentNullException(nameof(builder));
        return builder.AddHandler(CreateRegistration(pdfOptions), replaceExisting);
    }

    /// <summary>
    /// Unregisters PDF ingestion handler from <see cref="DocumentReader"/>.
    /// </summary>
    public static bool UnregisterPdfHandler() {
        return DocumentReader.UnregisterHandler(HandlerId);
    }

    private static ReaderHandlerRegistration CreateRegistration(ReaderPdfOptions? pdfOptions) {
        ReaderPdfOptions? registeredOptions = ReaderPdfOptionsCloner.CloneNullable(pdfOptions);
        return new ReaderHandlerRegistration {
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
                cancellationToken: ct),
            ReadDocumentPath = (path, readerOptions, ct) => DocumentReaderPdfExtensions.ReadPdfDocument(
                pdfPath: path,
                readerOptions: readerOptions,
                pdfOptions: ReaderPdfOptionsCloner.CloneNullable(registeredOptions),
                cancellationToken: ct),
            ReadDocumentStream = (stream, sourceName, readerOptions, ct) => DocumentReaderPdfExtensions.ReadPdfDocument(
                pdfStream: stream,
                sourceName: sourceName,
                readerOptions: readerOptions,
                pdfOptions: ReaderPdfOptionsCloner.CloneNullable(registeredOptions),
                cancellationToken: ct)
        };
    }
}
