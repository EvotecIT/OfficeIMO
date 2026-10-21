using System.Threading;

namespace OfficeIMO.Pdf;

public sealed partial class PdfDocumentPages {
    /// <summary>Deletes the selected pages with cancellation through selection, rewriting, and output adoption.</summary>
    public PdfDocument Delete(PdfPageSelection selection, CancellationToken cancellationToken) {
        Guard.NotNull(selection, nameof(selection));
        return Delete(selection, _document.ReadOptions, cancellationToken);
    }

    /// <summary>Deletes explicit pages while observing cancellation. The source document is unchanged.</summary>
    public PdfDocument Delete(CancellationToken cancellationToken, params int[] pageNumbers) =>
        _document.ApplyMutation(input => PdfPageEditor.DeletePagesWithReadOptions(input, _document.ReadOptions, cancellationToken, pageNumbers),
            cancellationToken: cancellationToken);

    /// <summary>Deletes pages resolved by a document-relative selector with cooperative cancellation.</summary>
    public PdfDocument Delete(PdfPageSelector selector, CancellationToken cancellationToken) =>
        Delete(ResolveSelector(selector, _document.ReadOptions, cancellationToken), _document.ReadOptions, cancellationToken);

    /// <summary>Deletes parsed page ranges with cooperative cancellation.</summary>
    public PdfDocument Delete(string pageRanges, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        return Delete(PdfPageSelection.Parse(pageRanges), _document.ReadOptions, cancellationToken);
    }

    /// <summary>Deletes an inclusive page range with cooperative cancellation.</summary>
    public PdfDocument Delete(PdfPageRange pageRange, CancellationToken cancellationToken) =>
        Delete(PdfPageSelection.FromRanges(pageRange), cancellationToken);

    /// <summary>Reports the delete operation while propagating cancellation as an exception.</summary>
    public PdfOperationResult<PdfDocument> DeleteResult(PdfPageSelection selection, CancellationToken cancellationToken, PdfLoadOptions? options = null) {
        PdfLoadOptions effective = options ?? _document.ReadOptions;
        return _document.TryMutationOperation("Delete pages", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.OrganizePages,
            () => Delete(selection, effective, cancellationToken), effective, cancellationToken: cancellationToken);
    }

    /// <summary>Reports the delete operation while propagating cancellation as an exception.</summary>
    public PdfOperationResult<PdfDocument> DeleteResult(PdfPageSelector selector, CancellationToken cancellationToken, PdfLoadOptions? options = null) {
        PdfLoadOptions effective = options ?? _document.ReadOptions;
        return _document.TryMutationOperation("Delete pages", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.OrganizePages,
            () => Delete(ResolveSelector(selector, effective, cancellationToken), effective, cancellationToken), effective, cancellationToken: cancellationToken);
    }

    /// <summary>Reports the delete operation while propagating cancellation as an exception.</summary>
    public PdfOperationResult<PdfDocument> DeleteResult(string pageRanges, CancellationToken cancellationToken, PdfLoadOptions? options = null) {
        PdfLoadOptions effective = options ?? _document.ReadOptions;
        return _document.TryMutationOperation("Delete pages", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.OrganizePages,
            () => Delete(PdfPageSelection.Parse(pageRanges), effective, cancellationToken), effective, cancellationToken: cancellationToken);
    }

    /// <summary>Reorders the selected pages with cancellation through selection, rewriting, and output adoption.</summary>
    public PdfDocument Reorder(PdfPageSelection selection, CancellationToken cancellationToken) {
        Guard.NotNull(selection, nameof(selection));
        return Reorder(selection, _document.ReadOptions, cancellationToken);
    }

    /// <summary>Reorders explicit pages while observing cancellation. The source document is unchanged.</summary>
    public PdfDocument Reorder(CancellationToken cancellationToken, params int[] pageNumbers) =>
        _document.ApplyMutation(input => PdfPageEditor.ReorderPagesWithReadOptions(input, _document.ReadOptions, cancellationToken, pageNumbers),
            cancellationToken: cancellationToken);

    /// <summary>Reorders pages resolved by a document-relative selector with cooperative cancellation.</summary>
    public PdfDocument Reorder(PdfPageSelector selector, CancellationToken cancellationToken) =>
        Reorder(ResolveSelector(selector, _document.ReadOptions, cancellationToken), _document.ReadOptions, cancellationToken);

    /// <summary>Reorders parsed page ranges with cooperative cancellation.</summary>
    public PdfDocument Reorder(string pageRanges, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        return Reorder(PdfPageSelection.Parse(pageRanges), _document.ReadOptions, cancellationToken);
    }

    /// <summary>Reports the reorder operation while propagating cancellation as an exception.</summary>
    public PdfOperationResult<PdfDocument> ReorderResult(PdfPageSelection selection, CancellationToken cancellationToken, PdfLoadOptions? options = null) {
        PdfLoadOptions effective = options ?? _document.ReadOptions;
        return _document.TryMutationOperation("Reorder pages", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.OrganizePages,
            () => Reorder(selection, effective, cancellationToken), effective, cancellationToken: cancellationToken);
    }

    /// <summary>Reports the reorder operation while propagating cancellation as an exception.</summary>
    public PdfOperationResult<PdfDocument> ReorderResult(PdfPageSelector selector, CancellationToken cancellationToken, PdfLoadOptions? options = null) {
        PdfLoadOptions effective = options ?? _document.ReadOptions;
        return _document.TryMutationOperation("Reorder pages", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.OrganizePages,
            () => Reorder(ResolveSelector(selector, effective, cancellationToken), effective, cancellationToken), effective, cancellationToken: cancellationToken);
    }

    /// <summary>Reports the reorder operation while propagating cancellation as an exception.</summary>
    public PdfOperationResult<PdfDocument> ReorderResult(string pageRanges, CancellationToken cancellationToken, PdfLoadOptions? options = null) {
        PdfLoadOptions effective = options ?? _document.ReadOptions;
        return _document.TryMutationOperation("Reorder pages", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.OrganizePages,
            () => Reorder(PdfPageSelection.Parse(pageRanges), effective, cancellationToken), effective, cancellationToken: cancellationToken);
    }

    /// <summary>Duplicates the selected pages with cancellation through selection, rewriting, and output adoption.</summary>
    public PdfDocument Duplicate(PdfPageSelection selection, CancellationToken cancellationToken) {
        Guard.NotNull(selection, nameof(selection));
        return Duplicate(selection, _document.ReadOptions, cancellationToken);
    }

    /// <summary>Duplicates explicit pages while observing cancellation. The source document is unchanged.</summary>
    public PdfDocument Duplicate(CancellationToken cancellationToken, params int[] pageNumbers) =>
        _document.ApplyMutation(input => PdfPageEditor.DuplicatePagesWithReadOptions(input, _document.ReadOptions, cancellationToken, pageNumbers),
            cancellationToken: cancellationToken);

    /// <summary>Duplicates pages resolved by a document-relative selector with cooperative cancellation.</summary>
    public PdfDocument Duplicate(PdfPageSelector selector, CancellationToken cancellationToken) =>
        Duplicate(ResolveSelector(selector, _document.ReadOptions, cancellationToken), _document.ReadOptions, cancellationToken);

    /// <summary>Duplicates parsed page ranges with cooperative cancellation.</summary>
    public PdfDocument Duplicate(string pageRanges, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        return Duplicate(PdfPageSelection.Parse(pageRanges), _document.ReadOptions, cancellationToken);
    }

    /// <summary>Duplicates an inclusive page range with cooperative cancellation.</summary>
    public PdfDocument Duplicate(PdfPageRange pageRange, CancellationToken cancellationToken) =>
        Duplicate(PdfPageSelection.FromRanges(pageRange), cancellationToken);

    /// <summary>Reports the duplicate operation while propagating cancellation as an exception.</summary>
    public PdfOperationResult<PdfDocument> DuplicateResult(PdfPageSelection selection, CancellationToken cancellationToken, PdfLoadOptions? options = null) {
        PdfLoadOptions effective = options ?? _document.ReadOptions;
        return _document.TryMutationOperation("Duplicate pages", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.OrganizePages,
            () => Duplicate(selection, effective, cancellationToken), effective, cancellationToken: cancellationToken);
    }

    /// <summary>Reports the duplicate operation while propagating cancellation as an exception.</summary>
    public PdfOperationResult<PdfDocument> DuplicateResult(PdfPageSelector selector, CancellationToken cancellationToken, PdfLoadOptions? options = null) {
        PdfLoadOptions effective = options ?? _document.ReadOptions;
        return _document.TryMutationOperation("Duplicate pages", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.OrganizePages,
            () => Duplicate(ResolveSelector(selector, effective, cancellationToken), effective, cancellationToken), effective, cancellationToken: cancellationToken);
    }

    /// <summary>Reports the duplicate operation while propagating cancellation as an exception.</summary>
    public PdfOperationResult<PdfDocument> DuplicateResult(string pageRanges, CancellationToken cancellationToken, PdfLoadOptions? options = null) {
        PdfLoadOptions effective = options ?? _document.ReadOptions;
        return _document.TryMutationOperation("Duplicate pages", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.OrganizePages,
            () => Duplicate(PdfPageSelection.Parse(pageRanges), effective, cancellationToken), effective, cancellationToken: cancellationToken);
    }

    /// <summary>Moves the selected pages with cancellation through selection, rewriting, and output adoption.</summary>
    public PdfDocument Move(int insertBeforePageNumber, PdfPageSelection selection, CancellationToken cancellationToken) {
        Guard.NotNull(selection, nameof(selection));
        return Move(insertBeforePageNumber, selection, _document.ReadOptions, cancellationToken);
    }

    /// <summary>Moves explicit pages while observing cancellation. The source document is unchanged.</summary>
    public PdfDocument Move(int insertBeforePageNumber, CancellationToken cancellationToken, params int[] pageNumbers) =>
        _document.ApplyMutation(input => PdfPageEditor.MovePagesWithReadOptions(input, insertBeforePageNumber, _document.ReadOptions, cancellationToken, pageNumbers),
            cancellationToken: cancellationToken);

    /// <summary>Moves pages resolved by a document-relative selector with cooperative cancellation.</summary>
    public PdfDocument Move(int insertBeforePageNumber, PdfPageSelector selector, CancellationToken cancellationToken) =>
        Move(insertBeforePageNumber, ResolveSelector(selector, _document.ReadOptions, cancellationToken), _document.ReadOptions, cancellationToken);

    /// <summary>Moves parsed page ranges with cooperative cancellation.</summary>
    public PdfDocument Move(int insertBeforePageNumber, string pageRanges, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        return Move(insertBeforePageNumber, PdfPageSelection.Parse(pageRanges), _document.ReadOptions, cancellationToken);
    }

    /// <summary>Moves an inclusive page range with cooperative cancellation.</summary>
    public PdfDocument Move(int insertBeforePageNumber, PdfPageRange pageRange, CancellationToken cancellationToken) =>
        Move(insertBeforePageNumber, PdfPageSelection.FromRanges(pageRange), cancellationToken);

    /// <summary>Reports the move operation while propagating cancellation as an exception.</summary>
    public PdfOperationResult<PdfDocument> MoveResult(int insertBeforePageNumber, PdfPageSelection selection, CancellationToken cancellationToken, PdfLoadOptions? options = null) {
        PdfLoadOptions effective = options ?? _document.ReadOptions;
        return _document.TryMutationOperation("Move pages", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.OrganizePages,
            () => Move(insertBeforePageNumber, selection, effective, cancellationToken), effective, cancellationToken: cancellationToken);
    }

    /// <summary>Reports the move operation while propagating cancellation as an exception.</summary>
    public PdfOperationResult<PdfDocument> MoveResult(int insertBeforePageNumber, PdfPageSelector selector, CancellationToken cancellationToken, PdfLoadOptions? options = null) {
        PdfLoadOptions effective = options ?? _document.ReadOptions;
        return _document.TryMutationOperation("Move pages", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.OrganizePages,
            () => Move(insertBeforePageNumber, ResolveSelector(selector, effective, cancellationToken), effective, cancellationToken), effective, cancellationToken: cancellationToken);
    }

    /// <summary>Reports the move operation while propagating cancellation as an exception.</summary>
    public PdfOperationResult<PdfDocument> MoveResult(int insertBeforePageNumber, string pageRanges, CancellationToken cancellationToken, PdfLoadOptions? options = null) {
        PdfLoadOptions effective = options ?? _document.ReadOptions;
        return _document.TryMutationOperation("Move pages", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.OrganizePages,
            () => Move(insertBeforePageNumber, PdfPageSelection.Parse(pageRanges), effective, cancellationToken), effective, cancellationToken: cancellationToken);
    }

    /// <summary>Rotates the selected pages with cancellation through selection, rewriting, and output adoption.</summary>
    public PdfDocument Rotate(int rotationDegrees, PdfPageSelection selection, CancellationToken cancellationToken) {
        Guard.NotNull(selection, nameof(selection));
        return Rotate(rotationDegrees, selection, _document.ReadOptions, cancellationToken);
    }

    /// <summary>Rotates explicit pages while observing cancellation. The source document is unchanged.</summary>
    public PdfDocument Rotate(int rotationDegrees, CancellationToken cancellationToken, params int[] pageNumbers) =>
        _document.ApplyMutation(input => PdfPageEditor.RotatePagesWithReadOptions(input, rotationDegrees, _document.ReadOptions, cancellationToken, pageNumbers),
            cancellationToken: cancellationToken);

    /// <summary>Rotates pages resolved by a document-relative selector with cooperative cancellation.</summary>
    public PdfDocument Rotate(int rotationDegrees, PdfPageSelector selector, CancellationToken cancellationToken) =>
        Rotate(rotationDegrees, ResolveSelector(selector, _document.ReadOptions, cancellationToken), _document.ReadOptions, cancellationToken);

    /// <summary>Rotates parsed page ranges with cooperative cancellation.</summary>
    public PdfDocument Rotate(int rotationDegrees, string pageRanges, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        return Rotate(rotationDegrees, PdfPageSelection.Parse(pageRanges), _document.ReadOptions, cancellationToken);
    }

    /// <summary>Rotates an inclusive page range with cooperative cancellation.</summary>
    public PdfDocument Rotate(int rotationDegrees, PdfPageRange pageRange, CancellationToken cancellationToken) =>
        Rotate(rotationDegrees, PdfPageSelection.FromRanges(pageRange), cancellationToken);

    /// <summary>Reports the rotate operation while propagating cancellation as an exception.</summary>
    public PdfOperationResult<PdfDocument> RotateResult(int rotationDegrees, PdfPageSelection selection, CancellationToken cancellationToken, PdfLoadOptions? options = null) {
        PdfLoadOptions effective = options ?? _document.ReadOptions;
        return _document.TryMutationOperation("Rotate pages", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.RotatePages,
            () => Rotate(rotationDegrees, selection, effective, cancellationToken), effective, cancellationToken: cancellationToken);
    }

    /// <summary>Reports the rotate operation while propagating cancellation as an exception.</summary>
    public PdfOperationResult<PdfDocument> RotateResult(int rotationDegrees, PdfPageSelector selector, CancellationToken cancellationToken, PdfLoadOptions? options = null) {
        PdfLoadOptions effective = options ?? _document.ReadOptions;
        return _document.TryMutationOperation("Rotate pages", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.RotatePages,
            () => Rotate(rotationDegrees, ResolveSelector(selector, effective, cancellationToken), effective, cancellationToken), effective, cancellationToken: cancellationToken);
    }

    /// <summary>Reports the rotate operation while propagating cancellation as an exception.</summary>
    public PdfOperationResult<PdfDocument> RotateResult(int rotationDegrees, string pageRanges, CancellationToken cancellationToken, PdfLoadOptions? options = null) {
        PdfLoadOptions effective = options ?? _document.ReadOptions;
        return _document.TryMutationOperation("Rotate pages", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.RotatePages,
            () => Rotate(rotationDegrees, PdfPageSelection.Parse(pageRanges), effective, cancellationToken), effective, cancellationToken: cancellationToken);
    }

    /// <summary>Applies page-attribute edits atomically while observing cancellation.</summary>
    public PdfDocument ApplyAttributes(CancellationToken cancellationToken, params PdfPageAttributeEdit[] edits) =>
        _document.ApplyMutation(input => PdfPageEditor.EditPageAttributes(input, edits, _document.ReadOptions, cancellationToken),
            cancellationToken: cancellationToken);

    /// <summary>Turns selected pages relative to their effective rotation with cooperative cancellation.</summary>
    public PdfDocument RotateBy(int rotationDegrees, PdfPageSelection selection, CancellationToken cancellationToken) =>
        ApplyAttributes(cancellationToken, PdfPageAttributeEdit.RotateBy(rotationDegrees, selection));

    /// <summary>Turns explicit pages relative to their effective rotation with cooperative cancellation.</summary>
    public PdfDocument RotateBy(int rotationDegrees, CancellationToken cancellationToken, params int[] pageNumbers) =>
        _document.ApplyMutation(input => PdfPageEditor.RotatePagesByWithReadOptions(input, rotationDegrees, _document.ReadOptions, cancellationToken, pageNumbers),
            cancellationToken: cancellationToken);

    /// <summary>Turns pages resolved by a document-relative selector with cooperative cancellation.</summary>
    public PdfDocument RotateBy(int rotationDegrees, PdfPageSelector selector, CancellationToken cancellationToken) =>
        RotateBy(rotationDegrees, ResolveSelector(selector, _document.ReadOptions, cancellationToken), cancellationToken);

    /// <summary>Turns pages resolved by a document-relative expression with cooperative cancellation.</summary>
    public PdfDocument RotateBy(int rotationDegrees, string selector, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        return RotateBy(rotationDegrees, PdfPageSelector.Parse(selector), cancellationToken);
    }

    /// <summary>Sets page boundary geometry with cooperative cancellation.</summary>
    public PdfDocument SetPageBox(PdfPageBoundaryBox box, double left, double bottom, double right, double top, CancellationToken cancellationToken, params int[] pageNumbers) =>
        _document.ApplyMutation(input => PdfPageEditor.SetPageBoxWithReadOptions(input, box, left, bottom, right, top, _document.ReadOptions, cancellationToken, pageNumbers),
            cancellationToken: cancellationToken);

    /// <summary>Resizes selected pages with cooperative cancellation.</summary>
    public PdfDocument Resize(PageSize pageSize, CancellationToken cancellationToken, params int[] pageNumbers) =>
        Resize(new PdfPageResizeOptions(pageSize) { CancellationToken = cancellationToken }, pageNumbers);

    /// <summary>Crops and translates selected pages while retaining the source read options and observing cancellation.</summary>
    public PdfDocument CropAndTranslate(double left, double bottom, double right, double top, CancellationToken cancellationToken, params int[] pageNumbers) =>
        _document.ApplyMutation(input => PdfPageEditor.CropAndTranslatePages(input, left, bottom, right, top, _document.ReadOptions, cancellationToken, pageNumbers),
            cancellationToken: cancellationToken);

    /// <summary>Reorders pages and reports object remapping with cooperative cancellation.</summary>
    public PdfPageRewriteResult ReorderWithMapping(CancellationToken cancellationToken, params int[] pageNumbers) {
        byte[] input = _document.GetBytesForOperation(cancellationToken);
        var result = PdfPageEditor.ReorderPagesWithMapping(input, _document.ReadOptions, pageNumbers, cancellationToken);
        PdfDocument rewritten = _document.WithBytes(input, result.Bytes, cancellationToken: cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        return new PdfPageRewriteResult(result.Bytes, result.Map, rewritten.ReadOptions);
    }
    /// <summary>Sets the media box while observing cancellation.</summary>
    public PdfDocument SetMediaBox(double left, double bottom, double right, double top, CancellationToken cancellationToken, params int[] pageNumbers) =>
        SetPageBox(PdfPageBoundaryBox.MediaBox, left, bottom, right, top, cancellationToken, pageNumbers);

    /// <summary>Sets the crop box while observing cancellation.</summary>
    public PdfDocument SetCropBox(double left, double bottom, double right, double top, CancellationToken cancellationToken, params int[] pageNumbers) =>
        SetPageBox(PdfPageBoundaryBox.CropBox, left, bottom, right, top, cancellationToken, pageNumbers);

    /// <summary>Sets the bleed box while observing cancellation.</summary>
    public PdfDocument SetBleedBox(double left, double bottom, double right, double top, CancellationToken cancellationToken, params int[] pageNumbers) =>
        SetPageBox(PdfPageBoundaryBox.BleedBox, left, bottom, right, top, cancellationToken, pageNumbers);

    /// <summary>Sets the trim box while observing cancellation.</summary>
    public PdfDocument SetTrimBox(double left, double bottom, double right, double top, CancellationToken cancellationToken, params int[] pageNumbers) =>
        SetPageBox(PdfPageBoundaryBox.TrimBox, left, bottom, right, top, cancellationToken, pageNumbers);

    /// <summary>Sets the art box while observing cancellation.</summary>
    public PdfDocument SetArtBox(double left, double bottom, double right, double top, CancellationToken cancellationToken, params int[] pageNumbers) =>
        SetPageBox(PdfPageBoundaryBox.ArtBox, left, bottom, right, top, cancellationToken, pageNumbers);

    /// <summary>Reports a relative page rotation while propagating cancellation as an exception.</summary>
    public PdfOperationResult<PdfDocument> RotateByResult(int rotationDegrees, PdfPageSelection selection, CancellationToken cancellationToken, PdfLoadOptions? options = null) {
        PdfLoadOptions effective = options ?? _document.ReadOptions;
        return _document.TryMutationOperation("Turn pages", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.RotatePages,
            () => _document.ApplyMutation(input => PdfPageEditor.EditPageAttributes(input,
                new[] { PdfPageAttributeEdit.RotateBy(rotationDegrees, selection) }, effective, cancellationToken),
                effective, cancellationToken: cancellationToken), effective, cancellationToken: cancellationToken);
    }

    /// <summary>Reports a relative page rotation while propagating cancellation as an exception.</summary>
    public PdfOperationResult<PdfDocument> RotateByResult(int rotationDegrees, PdfPageSelector selector, CancellationToken cancellationToken, PdfLoadOptions? options = null) {
        PdfLoadOptions effective = options ?? _document.ReadOptions;
        return _document.TryMutationOperation("Turn pages", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.RotatePages,
            () => _document.ApplyMutation(input => PdfPageEditor.EditPageAttributes(input,
                new[] { PdfPageAttributeEdit.RotateBy(rotationDegrees, ResolveSelector(selector, effective, cancellationToken)) }, effective, cancellationToken),
                effective, cancellationToken: cancellationToken), effective, cancellationToken: cancellationToken);
    }

    /// <summary>Reports a relative page rotation while propagating cancellation as an exception.</summary>
    public PdfOperationResult<PdfDocument> RotateByResult(int rotationDegrees, string selector, CancellationToken cancellationToken, PdfLoadOptions? options = null) {
        PdfLoadOptions effective = options ?? _document.ReadOptions;
        return _document.TryMutationOperation("Turn pages", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.RotatePages,
            () => _document.ApplyMutation(input => PdfPageEditor.EditPageAttributes(input,
                new[] { PdfPageAttributeEdit.RotateBy(rotationDegrees, ResolveSelector(PdfPageSelector.Parse(selector), effective, cancellationToken)) }, effective, cancellationToken),
                effective, cancellationToken: cancellationToken), effective, cancellationToken: cancellationToken);
    }

    /// <summary>Reports a page-box edit while propagating cancellation as an exception.</summary>
    public PdfOperationResult<PdfDocument> SetPageBoxResult(PdfPageBoundaryBox box, double left, double bottom, double right, double top,
        PdfPageSelection selection, CancellationToken cancellationToken, PdfLoadOptions? options = null) {
        PdfLoadOptions effective = options ?? _document.ReadOptions;
        return _document.TryMutationOperation("Set page box", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.ModifyPageAttributes,
            () => _document.ApplyMutation(input => PdfPageEditor.SetPageBoxWithReadOptions(input, box, left, bottom, right, top, effective,
                cancellationToken, selection.ToPageNumbers(GetSelectionPageCount(effective, cancellationToken), nameof(selection))),
                effective, cancellationToken: cancellationToken), effective, cancellationToken: cancellationToken);
    }
}
