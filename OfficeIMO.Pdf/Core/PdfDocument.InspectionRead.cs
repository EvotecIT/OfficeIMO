using System.Threading;

namespace OfficeIMO.Pdf;

public sealed partial class PdfDocument {
    /// <summary>
    /// Reads supported selected pages while retaining structural facts and explicit outcomes for unmappable
    /// text and unsupported content-stream filters. Existing <see cref="Read"/> remains fail-fast.
    /// Cancellation, resource admission, authentication, permissions, structural parsing, and unexpected
    /// stage failures remain terminal and are never converted to partial success.
    /// </summary>
    public PdfInspectionReadResult ReadForInspection(PdfReadOptions? options = null,
        CancellationToken cancellationToken = default) {
        PdfReadOptions effective = PdfReadOptions.Resolve(options).Clone();
        var snapshot = GetReadSnapshot(cancellationToken: cancellationToken);
        snapshot.Document.DemandContentExtraction("logical object");
        PdfDocumentInfo sourceInfo = PdfInspector.Inspect(snapshot.Bytes, snapshot.Document, cancellationToken);
        return PdfDocumentReadEngine.ReadForInspection(snapshot.Document, sourceInfo, effective, cancellationToken);
    }
}
