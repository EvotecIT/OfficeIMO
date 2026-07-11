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
}
