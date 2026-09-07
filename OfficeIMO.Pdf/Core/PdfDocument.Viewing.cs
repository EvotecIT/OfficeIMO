using System.Threading;

namespace OfficeIMO.Pdf;

public sealed partial class PdfDocument {
    /// <summary>
    /// Inspects authenticated page geometry for viewing without requiring content-copy permission.
    /// Full logical inspection is included only when extraction is authorized. This does not change mutation or extraction policy.
    /// </summary>
    public PdfDocumentViewInfo InspectForViewing(PdfLoadOptions? options = null, CancellationToken cancellationToken = default) {
        var snapshot = GetReadSnapshot(options, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        PdfReadDocument document = snapshot.Document;
        bool canExtractText = PdfPermissionAuthorization.CanExtractText(document.Security, snapshot.Options.PermissionPolicy);
        if (PdfPermissionAuthorization.CanExtractContent(document.Security, snapshot.Options.PermissionPolicy)) {
            PdfDocumentInfo content = PdfInspector.Inspect(snapshot.Bytes, document, cancellationToken);
            return new PdfDocumentViewInfo(content.Pages, document.Security, canExtractText, content);
        }
        var pages = new List<PdfPageInfo>(document.Pages.Count);
        for (int index = 0; index < document.Pages.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            PdfReadPage page = document.Pages[index];
            var size = page.GetPageSize();
            pages.Add(new PdfPageInfo(index + 1, size.Width, size.Height, page.GetRotationDegrees(), page.GetGeometry()));
        }
        return new PdfDocumentViewInfo(pages.AsReadOnly(), document.Security, canExtractText, null);
    }
}
