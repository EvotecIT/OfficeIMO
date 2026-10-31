using OfficeIMO.Pdf;

namespace OfficeIMO.Studio.Features.Workspace;

internal sealed partial class PdfWorkspace {
    /// <summary>Retains exact written-artifact evidence for workspace publication and source-change guards.</summary>
    private async Task<PdfSaveResult> SaveWorkspacePdfAsync(byte[] bytes, Stream destination, CancellationToken token) {
        using PdfDocument document = LoadDocument(bytes);
        return await document.SaveWithEvidenceAsync(destination, token).ConfigureAwait(false);
    }
}
