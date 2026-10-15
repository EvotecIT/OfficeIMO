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
    /// <summary>Projects a completed native inspection while preserving every selected page occurrence and failure.</summary>
    /// <param name="inspection">Caller-owned inspection result, including its original selection outcomes.</param>
    /// <param name="sourceName">Logical source name; no file is opened at this path.</param>
    /// <param name="readerOptions">Projection and output budgets. A null inspection policy uses the default evidence budget.</param>
    /// <param name="pdfOptions">Projection settings. Select pages when creating the native inspection; reselection here is rejected.</param>
    /// <param name="cancellationToken">Cancellation observed during projection.</param>
    /// <returns>A Reader result with native occurrence identities and located failure evidence.</returns>
    /// <remarks>This does not reload the source, repeat native parsing, or reassess its original load limits.</remarks>
    public static OfficeDocumentReadResult ReadDocument(
        this PdfInspectionReadResult inspection,
        string sourceName = "document.pdf",
        ReaderOptions? readerOptions = null,
        ReaderPdfOptions? pdfOptions = null,
        CancellationToken cancellationToken = default) {
        if (inspection == null) throw new ArgumentNullException(nameof(inspection));
        if (sourceName == null) throw new ArgumentNullException(nameof(sourceName));
        cancellationToken.ThrowIfCancellationRequested();
        ReaderOptions effective = DocumentReaderEngine.NormalizeOptions(readerOptions);
        effective.Inspection ??= new ReaderInspectionOptions();
        using var scope = ReaderReadScope.Enter(effective);
        OfficeDocumentReadResult result = PdfReaderAdapter.ReadDocument(inspection, sourceName, effective, pdfOptions, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        return ReaderReadScope.Complete(result);
    }

}
