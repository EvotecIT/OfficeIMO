using OfficeIMO.Pdf;

namespace OfficeIMO.Reader.Pdf;

/// <summary>Projects an already loaded native PDF result into the normalized Reader model.</summary>
public static class PdfDocumentReaderExtensions {
    /// <summary>Reads the supplied logical pages without reopening, parsing, or rendering the PDF.</summary>
    /// <param name="document">Caller-owned native logical result to project.</param>
    /// <param name="sourceName">Logical source name used in locations; no file is opened at this path.</param>
    /// <param name="readerOptions">Reader projection, inspection, and output budget settings. Null uses the full projection.</param>
    /// <param name="pdfOptions">PDF projection settings, including page selection and chunk grouping.</param>
    /// <param name="cancellationToken">Cancellation observed during projection.</param>
    /// <returns>A normalized Reader result for the supplied logical source.</returns>
    /// <remarks>Native parsing, authentication, input limits, and source-page completion remain the responsibility
    /// of the caller's native load. A reduced logical result cannot recover unrequested source content.</remarks>
    public static OfficeDocumentReadResult ReadDocument(
        this PdfDocumentReadResult document,
        string sourceName = "document.pdf",
        ReaderOptions? readerOptions = null,
        ReaderPdfOptions? pdfOptions = null,
        CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (sourceName == null) throw new ArgumentNullException(nameof(sourceName));
        cancellationToken.ThrowIfCancellationRequested();
        ReaderOptions effective = DocumentReaderEngine.NormalizeOptions(readerOptions);
        using var scope = ReaderReadScope.Enter(effective);
        OfficeDocumentReadResult result = PdfReaderAdapter.ReadDocument(
            document, sourceName, effective, pdfOptions, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        return ReaderReadScope.Complete(result);
    }
}
