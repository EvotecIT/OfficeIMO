using System.Threading;

namespace OfficeIMO.Pdf;

public sealed partial class PdfDocument {
    /// <summary>
    /// Reads selected pages while retaining structural facts and explicit outcomes for unmappable text,
    /// unsupported content-stream filters, and opt-in recovery from damaged page streams.
    /// <see cref="Read"/> remains fail-fast.
    /// Pass <see cref="PdfLoadOptions.RecoverDecodableTextForInspection"/> to retain readable text on
    /// a page with a font-mapping failure when using the built-in glyph decoder, or from leading page
    /// content streams before a supported page content filter reports damaged data.
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
