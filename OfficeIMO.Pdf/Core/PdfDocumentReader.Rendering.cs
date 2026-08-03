using System.Threading;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

public sealed partial class PdfDocumentReader {
    /// <summary>
    /// Exports all pages or a caller-ordered selection through the shared image-export contract.
    /// </summary>
    public IReadOnlyList<OfficeImageExportResult> ExportImages(
        OfficeImageExportFormat format,
        PdfImageExportOptions? options = null,
        PdfPageSelection? selection = null,
        PdfReadOptions? readOptions = null,
        CancellationToken cancellationToken = default) {
        return PdfImageExportEngine.Export(
            token => {
                token.ThrowIfCancellationRequested();
                PdfReadDocument document = ReadDocument(readOptions);
                token.ThrowIfCancellationRequested();
                return document;
            },
            format,
            options?.Clone() ?? new PdfImageExportOptions(),
            _ => selection,
            initialDiagnostics: null,
            cancellationToken);
    }

    /// <summary>Renders all pages or a caller-ordered selection through the managed batch renderer.</summary>
    public IReadOnlyList<PdfPageRenderResult> RenderPages(
        PdfPageSelection? selection = null,
        PdfPageRenderOptions? options = null,
        PdfReadOptions? readOptions = null,
        CancellationToken cancellationToken = default) {
        return PdfPageImageRenderer.RenderPages(_document.GetBytesForOperation, selection, options, ResolveReadOptions(readOptions), cancellationToken);
    }

    /// <summary>Renders parsed page ranges such as <c>1-3,5</c> through the managed batch renderer.</summary>
    public IReadOnlyList<PdfPageRenderResult> RenderPages(
        string pageRanges,
        PdfPageRenderOptions? options = null,
        PdfReadOptions? readOptions = null,
        CancellationToken cancellationToken = default) {
        return PdfPageImageRenderer.RenderPages(_document.GetBytesForOperation, pageRanges, options, ResolveReadOptions(readOptions), cancellationToken);
    }
}
