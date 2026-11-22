namespace OfficeIMO.Pdf;

public sealed partial class PdfDocument {
#pragma warning disable RS0026, RS0027 // The stream, explicit-window and path inputs have distinct types.
    /// <summary>Loads a PDF from the current position of a readable seekable stream.</summary>
    /// <remarks>
    /// Keep the stream available and unchanged until the document is disposed. Each operation preserves its
    /// position. The caller owns the input stream. Initial admission scans the whole PDF with bounded reads;
    /// encoded streams are fetched as needed later. Use a separate stream for concurrent caller I/O.
    /// </remarks>
    public static PdfDocument LoadSeekable(Stream stream, PdfLoadOptions? loadOptions = null,
        CancellationToken cancellationToken = default) =>
        new PdfDocument(PdfDocumentSource.FromSeekableStream(stream, length: null, loadOptions, cancellationToken),
            cancellationToken);

    /// <summary>Loads a PDF stored in the specified stream window starting at the current position.</summary>
    /// <param name="stream">A readable seekable stream kept available and unchanged while the document is used.</param>
    /// <param name="length">Length of the complete PDF window in bytes.</param>
    /// <param name="loadOptions">Optional reader settings and limits.</param>
    /// <param name="cancellationToken">Cancels the initial read.</param>
    public static PdfDocument LoadSeekable(Stream stream, long length, PdfLoadOptions? loadOptions = null,
        CancellationToken cancellationToken = default) =>
        new PdfDocument(PdfDocumentSource.FromSeekableStream(stream, length, loadOptions, cancellationToken),
            cancellationToken);

    /// <summary>Loads a PDF for bounded reads when the system can exclude input writes.</summary>
    /// <remarks>The document holds a file it can protect against writes open until disposal. Other inputs use a bounded snapshot.</remarks>
    public static PdfDocument LoadSeekable(string path, PdfLoadOptions? loadOptions = null,
        CancellationToken cancellationToken = default) =>
        new PdfDocument(PdfDocumentSource.FromSeekablePath(path, loadOptions, cancellationToken), cancellationToken) {
            _filePath = System.IO.Path.GetFullPath(path)
        };
#pragma warning restore RS0026, RS0027
}
