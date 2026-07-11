namespace OfficeIMO.Pdf;

public sealed partial class PdfDocumentReader {
    /// <summary>Builds text-selection and interactive hit regions for one page in visual top-left coordinates.</summary>
    public PdfPageInteractionMap Interactions(
        int pageNumber,
        PdfPageInteractionOptions? interactionOptions = null,
        PdfReadOptions? readOptions = null) {
        return PdfPageInteractionMap.Create(
            _document.Snapshot(),
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
            _document.Snapshot(),
            pageNumber,
            overlayOptions,
            layoutOptions,
            ResolveReadOptions(readOptions));
    }
}
