using System.Threading;

namespace OfficeIMO.Pdf;

public sealed partial class PdfDocumentReader {
    /// <summary>Renders all pages or a caller-ordered selection through the managed batch renderer.</summary>
    public IReadOnlyList<PdfPageRenderResult> RenderPages(
        PdfPageSelection? selection = null,
        PdfPageRenderOptions? options = null,
        PdfReadOptions? readOptions = null,
        CancellationToken cancellationToken = default) {
        return PdfPageImageRenderer.RenderPages(_document.GetBytesForOperation(), selection, options, ResolveReadOptions(readOptions), cancellationToken);
    }

    /// <summary>Renders parsed page ranges such as <c>1-3,5</c> through the managed batch renderer.</summary>
    public IReadOnlyList<PdfPageRenderResult> RenderPages(
        string pageRanges,
        PdfPageRenderOptions? options = null,
        PdfReadOptions? readOptions = null,
        CancellationToken cancellationToken = default) {
        return PdfPageImageRenderer.RenderPages(_document.GetBytesForOperation(), pageRanges, options, ResolveReadOptions(readOptions), cancellationToken);
    }
}
