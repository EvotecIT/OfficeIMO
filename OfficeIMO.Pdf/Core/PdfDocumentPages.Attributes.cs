namespace OfficeIMO.Pdf;

public sealed partial class PdfDocumentPages {
    /// <summary>Applies ordered page-attribute edits atomically through one output revision or rewrite.</summary>
    public PdfDocument ApplyAttributes(params PdfPageAttributeEdit[] edits) =>
        _document.ApplyMutation(input => PdfPageEditor.EditPageAttributes(input, edits, _document.ReadOptions));

    /// <summary>Applies an atomic page-attribute batch and reports the route, reasons and blockers.</summary>
    /// <exception cref="OperationCanceledException">The requested cancellation interrupts planning, editing or output adoption.</exception>
    internal PdfOperationResult<PdfDocument> ApplyAttributesResult(IReadOnlyList<PdfPageAttributeEdit> edits,
        PdfLoadOptions? options = null, System.Threading.CancellationToken cancellationToken = default) {
        Guard.NotNull(edits, nameof(edits));
        PdfLoadOptions effective = options ?? _document.ReadOptions;
        return _document.TryMutationOperation("Edit page attributes", PdfPreflightCapability.ManipulatePages,
            PdfPageEditor.GetAttributeOperation(edits), () => _document.ApplyMutation(
                input => PdfPageEditor.EditPageAttributes(input, edits, effective, cancellationToken), effective,
                operationName: nameof(ApplyAttributes), cancellationToken: cancellationToken), effective,
            cancellationToken: cancellationToken);
    }

    /// <summary>
    /// Turns selected pages by a multiple of 90 degrees relative to each page's effective rotation.
    /// Positive values turn clockwise. If no page numbers are supplied, all pages are turned.
    /// Mixed inherited rotations are resolved independently in one mutation.
    /// </summary>
    public PdfDocument RotateBy(int rotationDegrees, params int[] pageNumbers) =>
        _document.ApplyMutation(input => PdfPageEditor.RotatePagesByWithReadOptions(
            input, rotationDegrees, _document.ReadOptions, pageNumbers));

    /// <summary>Turns the selected pages relative to their effective rotation.</summary>
    public PdfDocument RotateBy(int rotationDegrees, PdfPageSelection selection) =>
        ApplyAttributes(PdfPageAttributeEdit.RotateBy(rotationDegrees, selection ?? throw new ArgumentNullException(nameof(selection))));

    /// <summary>Turns pages resolved by a document-relative selector.</summary>
    public PdfDocument RotateBy(int rotationDegrees, PdfPageSelector selector) =>
        RotateBy(rotationDegrees, ResolveSelector(selector, _document.ReadOptions));

    /// <summary>Turns pages resolved by a document-relative selection expression.</summary>
    public PdfDocument RotateBy(int rotationDegrees, string selector) => RotateBy(rotationDegrees, PdfPageSelector.Parse(selector));

    /// <summary>Turns selected pages and reports the mutation route and preservation restrictions.</summary>
    internal PdfOperationResult<PdfDocument> RotateByResult(int rotationDegrees, PdfPageSelection selection, PdfLoadOptions? options = null) =>
        ApplyAttributesResult(new[] { PdfPageAttributeEdit.RotateBy(rotationDegrees,
            selection ?? throw new ArgumentNullException(nameof(selection))) }, options);

    /// <summary>Turns pages resolved by a document-relative selector and reports the mutation route.</summary>
    internal PdfOperationResult<PdfDocument> RotateByResult(int rotationDegrees, PdfPageSelector selector, PdfLoadOptions? options = null) {
        Guard.NotNull(selector, nameof(selector));
        return _document.TryMutationOperation("Turn pages", PdfPreflightCapability.ManipulatePages,
            PdfMutationOperation.RotatePages, () => _document.ApplyMutation(
                input => PdfPageEditor.EditPageAttributes(input, new[] { PdfPageAttributeEdit.RotateBy(rotationDegrees,
                    ResolveSelector(selector, options ?? _document.ReadOptions)) }, options ?? _document.ReadOptions), options ?? _document.ReadOptions,
                operationName: nameof(RotateBy)), options);
    }

    /// <summary>Turns pages resolved by a document-relative expression and reports the mutation route.</summary>
    internal PdfOperationResult<PdfDocument> RotateByResult(int rotationDegrees, string selector, PdfLoadOptions? options = null) =>
        _document.TryMutationOperation("Turn pages", PdfPreflightCapability.ManipulatePages,
            PdfMutationOperation.RotatePages, () => _document.ApplyMutation(
                input => PdfPageEditor.EditPageAttributes(input, new[] { PdfPageAttributeEdit.RotateBy(rotationDegrees,
                    ResolveSelector(PdfPageSelector.Parse(selector), options ?? _document.ReadOptions)) }, options ?? _document.ReadOptions), options ?? _document.ReadOptions,
                operationName: nameof(RotateBy)), options);

    /// <summary>
    /// Sets a page boundary box and returns the chosen mutation route, its reasons, and any blockers.
    /// This changes page geometry without translating or removing source content.
    /// </summary>
    internal PdfOperationResult<PdfDocument> SetPageBoxResult(
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
                selection.ToPageNumbers(GetSelectionPageCount(effectiveOptions), nameof(selection))), effectiveOptions,
                operationName: nameof(SetPageBox)),
            options: effectiveOptions);
    }
}
