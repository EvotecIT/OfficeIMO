namespace OfficeIMO.Pdf;

public sealed partial class PdfDocumentPages {
    /// <summary>
    /// Turns selected pages by a multiple of 90 degrees relative to each page's effective rotation.
    /// Positive values turn clockwise. If no page numbers are supplied, all pages are turned.
    /// Mixed inherited rotations are resolved independently in one mutation.
    /// </summary>
    public PdfDocument RotateBy(int rotationDegrees, params int[] pageNumbers) =>
        _document.ApplyMutation(input => PdfPageEditor.RotatePagesByWithReadOptions(
            input, rotationDegrees, _document.ReadOptions, pageNumbers));

    /// <summary>
    /// Sets a page boundary box and returns the chosen mutation route, its reasons, and any blockers.
    /// This changes page geometry without translating or removing source content.
    /// </summary>
    public PdfOperationResult<PdfDocument> SetPageBoxResult(
        PdfPageBoundaryBox box,
        double left,
        double bottom,
        double right,
        double top,
        PdfPageSelection selection,
        PdfLoadOptions? options = null) {
        Guard.NotNull(selection, nameof(selection));
        PdfLoadOptions effectiveOptions = options ?? _document.ReadOptions;
        return _document.TryMutationOperation(
            "Set page box", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.ModifyPageAttributes,
            () => _document.ApplyMutation(input => PdfPageEditor.SetPageBoxWithReadOptions(
                input, box, left, bottom, right, top, effectiveOptions,
                selection.ToPageNumbers(_document.Inspect(effectiveOptions).PageCount, nameof(selection))), effectiveOptions,
                operationName: nameof(SetPageBox)),
            options: effectiveOptions);
    }
}
