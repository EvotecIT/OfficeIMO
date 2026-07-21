namespace OfficeIMO.Pdf;

public sealed partial class PdfDocumentPages {
    /// <summary>Creates a new PDF containing pages resolved from a document-relative selector.</summary>
    public PdfDocument Extract(PdfPageSelector selector) {
        return Extract(ResolveSelector(selector, _document.ReadOptions), _document.ReadOptions);
    }

    /// <summary>Attempts to extract pages resolved from a document-relative selector.</summary>
    public PdfOperationResult<PdfDocument> TryExtract(PdfPageSelector selector, PdfReadOptions? options = null) {
        Guard.NotNull(selector, nameof(selector));
        return TryPageExtractionOperation("Extract pages", effectiveOptions => Extract(ResolveSelector(selector, effectiveOptions), effectiveOptions), options);
    }

    /// <summary>Creates a new PDF with pages resolved from a document-relative selector deleted.</summary>
    public PdfDocument Delete(PdfPageSelector selector) {
        return Delete(ResolveSelector(selector, _document.ReadOptions));
    }

    /// <summary>Attempts to delete pages resolved from a document-relative selector.</summary>
    public PdfOperationResult<PdfDocument> TryDelete(PdfPageSelector selector, PdfReadOptions? options = null) {
        Guard.NotNull(selector, nameof(selector));
        PdfReadOptions? effectiveOptions = options ?? _document.ReadOptions;
        return _document.TryMutationOperation(
            "Delete pages",
            PdfPreflightCapability.ManipulatePages,
            PdfMutationOperation.ModifyPageTree,
            () => Delete(ResolveSelector(selector, effectiveOptions), effectiveOptions),
            options);
    }

    /// <summary>Creates a new PDF with pages copied in the order resolved by a document-relative selector.</summary>
    public PdfDocument Reorder(PdfPageSelector selector) {
        return Reorder(ResolveSelector(selector, _document.ReadOptions));
    }

    /// <summary>Attempts to reorder pages resolved from a document-relative selector.</summary>
    public PdfOperationResult<PdfDocument> TryReorder(PdfPageSelector selector, PdfReadOptions? options = null) {
        Guard.NotNull(selector, nameof(selector));
        PdfReadOptions? effectiveOptions = options ?? _document.ReadOptions;
        return _document.TryMutationOperation(
            "Reorder pages",
            PdfPreflightCapability.ManipulatePages,
            PdfMutationOperation.ModifyPageTree,
            () => Reorder(ResolveSelector(selector, effectiveOptions), effectiveOptions),
            options);
    }

    /// <summary>Creates a new PDF with pages resolved from a document-relative selector duplicated.</summary>
    public PdfDocument Duplicate(PdfPageSelector selector) {
        return Duplicate(ResolveSelector(selector, _document.ReadOptions));
    }

    /// <summary>Attempts to duplicate pages resolved from a document-relative selector.</summary>
    public PdfOperationResult<PdfDocument> TryDuplicate(PdfPageSelector selector, PdfReadOptions? options = null) {
        Guard.NotNull(selector, nameof(selector));
        PdfReadOptions? effectiveOptions = options ?? _document.ReadOptions;
        return _document.TryMutationOperation(
            "Duplicate pages",
            PdfPreflightCapability.ManipulatePages,
            PdfMutationOperation.ModifyPageTree,
            () => Duplicate(ResolveSelector(selector, effectiveOptions), effectiveOptions),
            options);
    }

    /// <summary>Moves pages resolved from a document-relative selector before a one-based page number.</summary>
    public PdfDocument Move(int insertBeforePageNumber, PdfPageSelector selector) {
        return Move(insertBeforePageNumber, ResolveSelector(selector, _document.ReadOptions));
    }

    /// <summary>Attempts to move pages resolved from a document-relative selector.</summary>
    public PdfOperationResult<PdfDocument> TryMove(int insertBeforePageNumber, PdfPageSelector selector, PdfReadOptions? options = null) {
        Guard.NotNull(selector, nameof(selector));
        PdfReadOptions? effectiveOptions = options ?? _document.ReadOptions;
        return _document.TryMutationOperation(
            "Move pages",
            PdfPreflightCapability.ManipulatePages,
            PdfMutationOperation.ModifyPageTree,
            () => Move(insertBeforePageNumber, ResolveSelector(selector, effectiveOptions), effectiveOptions),
            options);
    }

    /// <summary>Rotates pages resolved from a document-relative selector.</summary>
    public PdfDocument Rotate(int rotationDegrees, PdfPageSelector selector) {
        return Rotate(rotationDegrees, ResolveSelector(selector, _document.ReadOptions));
    }

    /// <summary>Attempts to rotate pages resolved from a document-relative selector.</summary>
    public PdfOperationResult<PdfDocument> TryRotate(int rotationDegrees, PdfPageSelector selector, PdfReadOptions? options = null) {
        Guard.NotNull(selector, nameof(selector));
        PdfReadOptions? effectiveOptions = options ?? _document.ReadOptions;
        return _document.TryMutationOperation(
            "Rotate pages",
            PdfPreflightCapability.ManipulatePages,
            PdfMutationOperation.ModifyPageTree,
            () => Rotate(rotationDegrees, ResolveSelector(selector, effectiveOptions), effectiveOptions),
            options);
    }

    private PdfPageSelection ResolveSelector(PdfPageSelector selector, PdfReadOptions? options) {
        Guard.NotNull(selector, nameof(selector));
        int pageCount = _document.Inspect(options).PageCount;
        if (pageCount < 1) {
            throw new InvalidOperationException("PDF does not contain any readable pages.");
        }

        return selector.ResolveSelection(pageCount);
    }
}
