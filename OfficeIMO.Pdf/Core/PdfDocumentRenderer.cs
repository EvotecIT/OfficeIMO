using System.Threading;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>
/// Managed page rendering and drawing projection for a loaded PDF document.
/// </summary>
public sealed class PdfDocumentRenderer {
    private readonly PdfDocument _document;

    internal PdfDocumentRenderer(PdfDocument document) {
        _document = document;
    }

    /// <summary>Projects a one-based PDF page into the shared editable drawing scene.</summary>
    public OfficeDrawing Drawing(int pageNumber) => _document.Reader.Drawing(pageNumber);

    /// <summary>Creates a drawing overlay that visualizes recovered page interactions and layout regions.</summary>
    public OfficeDrawing LayoutDebugOverlay(
        int pageNumber,
        PdfLayoutDebugOverlayOptions? options = null,
        PdfTextLayoutOptions? layoutOptions = null,
        PdfLoadOptions? loadOptions = null) =>
        _document.Reader.LayoutDebugOverlay(pageNumber, options, layoutOptions, loadOptions);

    /// <summary>Returns managed-renderer capability diagnostics for a one-based page.</summary>
    public IReadOnlyList<PdfRenderCapabilityDiagnostic> CapabilityDiagnostics(int pageNumber) =>
        _document.Reader.RenderCapabilityDiagnostics(pageNumber);

    /// <summary>Exports all pages or a caller-ordered selection through the shared image-export contract.</summary>
    public IReadOnlyList<OfficeImageExportResult> ExportImages(
        OfficeImageExportFormat format,
        PdfImageExportOptions? options = null,
        PdfPageSelection? selection = null,
        CancellationToken cancellationToken = default) =>
        _document.Reader.ExportImages(format, options, selection, cancellationToken: cancellationToken);

    /// <summary>Renders all pages or a caller-ordered selection through the managed batch renderer.</summary>
    public IReadOnlyList<PdfPageRenderResult> Pages(
        PdfPageSelection? selection = null,
        PdfPageRenderOptions? options = null,
        CancellationToken cancellationToken = default) =>
        _document.Reader.RenderPages(selection, options, cancellationToken: cancellationToken);

    /// <summary>Renders parsed page ranges such as <c>1-3,5</c> through the managed batch renderer.</summary>
    public IReadOnlyList<PdfPageRenderResult> Pages(
        string pageRanges,
        PdfPageRenderOptions? options = null,
        CancellationToken cancellationToken = default) =>
        _document.Reader.RenderPages(pageRanges, options, cancellationToken: cancellationToken);
}
