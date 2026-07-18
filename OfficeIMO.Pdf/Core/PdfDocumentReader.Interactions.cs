namespace OfficeIMO.Pdf;

public sealed partial class PdfDocumentReader {
    /// <summary>Builds text-selection and interactive hit regions for one page in visual top-left coordinates.</summary>
    public PdfPageInteractionMap Interactions(
        int pageNumber,
        PdfPageInteractionOptions? interactionOptions = null,
        PdfReadOptions? readOptions = null) {
        return PdfPageInteractionMap.Create(
            _document.GetBytesForOperation(),
            pageNumber,
            interactionOptions,
            ResolveReadOptions(readOptions));
    }

    /// <summary>Creates a shared-Drawing word, line, region, and reading-order diagnostic overlay.</summary>
    public OfficeIMO.Drawing.OfficeDrawing LayoutDebugOverlay(
        int pageNumber,
        PdfLayoutDebugOverlayOptions? overlayOptions = null,
        PdfTextLayoutOptions? layoutOptions = null,
        PdfReadOptions? readOptions = null) {
        return PdfLayoutDebugOverlay.CreateDrawing(
            _document.GetBytesForOperation(),
            pageNumber,
            overlayOptions,
            layoutOptions,
            ResolveReadOptions(readOptions));
    }

    /// <summary>Exports the first-party logical model to JSON, Markdown, ALTO XML, hOCR, or PAGE XML.</summary>
    public string ExportStructured(
        PdfStructuredExportFormat format,
        PdfTextLayoutOptions? layoutOptions = null) =>
        ExportStructured(format, layoutOptions, readOptions: null);

    /// <summary>Exports the first-party logical model with explicit read limits or credentials.</summary>
    public string ExportStructured(
        PdfStructuredExportFormat format,
        PdfTextLayoutOptions? layoutOptions,
        PdfReadOptions? readOptions) {
        return PdfStructuredExportEngine.Export(_document.GetBytesForOperation(), format, layoutOptions, ResolveReadOptions(readOptions));
    }
}
