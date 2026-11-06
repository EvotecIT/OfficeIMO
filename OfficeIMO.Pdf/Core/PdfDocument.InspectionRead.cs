using System.Threading;

namespace OfficeIMO.Pdf;

public sealed partial class PdfDocument {
    /// <summary>
    /// Reads supported selected pages while retaining structural facts and explicit outcomes for unmappable
    /// text and unsupported content-stream filters. Existing <see cref="Read"/> remains fail-fast.
    /// Pass <see cref="PdfLoadOptions.RecoverDecodableTextForInspection"/> to retain readable text on
    /// a page with a font-mapping failure when using the built-in glyph decoder.
    /// Cancellation, resource admission, authentication, permissions, structural parsing, and unexpected
    /// stage failures remain terminal and are never converted to partial success.
    /// </summary>
    public PdfInspectionReadResult ReadForInspection(PdfLoadOptions? options = null,
        CancellationToken cancellationToken = default) {
        PdfLoadOptions effective = PdfLoadOptions.Resolve(options).Clone();
        var snapshot = GetReadSnapshot(cancellationToken: cancellationToken);
        snapshot.Document.DemandContentExtraction("logical object");
        PdfDocumentInfo sourceInfo = PdfInspector.Inspect(snapshot.Bytes, snapshot.Document, cancellationToken);
        bool recoverDecodableText = options is null
            ? snapshot.Options.RecoverDecodableTextForInspection
            : effective.RecoverDecodableTextForInspection;
        return PdfDocumentReadEngine.ReadForInspection(snapshot.Document, sourceInfo, effective, recoverDecodableText, cancellationToken);
    }
}
