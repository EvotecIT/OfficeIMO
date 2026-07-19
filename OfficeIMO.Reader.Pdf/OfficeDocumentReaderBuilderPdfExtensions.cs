namespace OfficeIMO.Reader.Pdf;

/// <summary>
/// Adds PDF support to <see cref="OfficeDocumentReaderBuilder"/>.
/// </summary>
public static class OfficeDocumentReaderBuilderPdfExtensions {
    /// <summary>
    /// Stable handler identifier for PDF adapter registration.
    /// </summary>
    public const string HandlerId = "officeimo.reader.pdf";

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

    private static ReaderHandlerRegistration CreateRegistration(ReaderPdfOptions? pdfOptions) {
        ReaderPdfOptions? registeredOptions = ReaderPdfOptionsCloner.CloneNullable(pdfOptions);
        return new ReaderHandlerRegistration {
            Origin = ReaderHandlerOrigin.OfficeIMO,
            Id = HandlerId,
            DisplayName = "PDF Reader Adapter",
            Description = "Modular PDF adapter using OfficeIMO.Pdf logical read model.",
            Kind = ReaderInputKind.Pdf,
            Extensions = new[] { ".pdf" },
            ReadPath = (path, readerOptions, ct) => PdfReaderAdapter.Read(
                pdfPath: path,
                readerOptions: readerOptions,
                pdfOptions: ReaderPdfOptionsCloner.CloneNullable(registeredOptions),
                cancellationToken: ct),
            ReadStream = (stream, sourceName, readerOptions, ct) => PdfReaderAdapter.Read(
                pdfStream: stream,
                sourceName: sourceName,
                readerOptions: readerOptions,
                pdfOptions: ReaderPdfOptionsCloner.CloneNullable(registeredOptions),
                cancellationToken: ct),
            ReadDocumentPath = (path, readerOptions, ct) => PdfReaderAdapter.ReadDocument(
                pdfPath: path,
                readerOptions: readerOptions,
                pdfOptions: ReaderPdfOptionsCloner.CloneNullable(registeredOptions),
                cancellationToken: ct),
            ReadDocumentStream = (stream, sourceName, readerOptions, ct) => PdfReaderAdapter.ReadDocument(
                pdfStream: stream,
                sourceName: sourceName,
                readerOptions: readerOptions,
                pdfOptions: ReaderPdfOptionsCloner.CloneNullable(registeredOptions),
                cancellationToken: ct)
        };
    }
}
