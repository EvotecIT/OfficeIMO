using System.Threading;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

public sealed partial class PdfDocumentReader {
    /// <summary>Projects a one-based PDF page into the shared editable drawing scene.</summary>
    public OfficeDrawing Drawing(int pageNumber, PdfReadOptions? readOptions = null) {
        PdfReadDocument document = ReadDocument(readOptions);
        if (pageNumber <= 0 || pageNumber > document.Pages.Count) {
            throw new ArgumentOutOfRangeException(nameof(pageNumber), pageNumber, "Page number must refer to an existing one-based PDF page.");
        }
        return document.Pages[pageNumber - 1].ToDrawing();
    }

    /// <summary>Returns managed-renderer capability diagnostics for a one-based PDF page.</summary>
    public IReadOnlyList<PdfRenderCapabilityDiagnostic> RenderCapabilityDiagnostics(
        int pageNumber,
        PdfReadOptions? readOptions = null) {
        PdfReadDocument document = ReadDocument(readOptions);
        if (pageNumber <= 0 || pageNumber > document.Pages.Count) {
            throw new ArgumentOutOfRangeException(nameof(pageNumber), pageNumber, "Page number must refer to an existing one-based PDF page.");
        }
        return document.Pages[pageNumber - 1].GetRenderCapabilityDiagnostics();
    }

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
