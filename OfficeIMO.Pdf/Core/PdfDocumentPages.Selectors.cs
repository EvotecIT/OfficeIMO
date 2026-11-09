using System.Threading;

namespace OfficeIMO.Pdf;

public sealed partial class PdfDocumentPages {
    /// <summary>Creates a new PDF containing pages resolved from a document-relative selector.</summary>
    public PdfDocument Extract(OfficePageSelector selector) {
        return Extract(ResolveSelector(selector, _document.ReadOptions), _document.ReadOptions);
    }

    /// <summary>Attempts to extract pages resolved from a document-relative selector.</summary>
    internal PdfOperationResult<PdfDocument> ExtractResult(OfficePageSelector selector, PdfLoadOptions? options = null) {
        Guard.NotNull(selector, nameof(selector));
        return TryPageExtractionOperation("Extract pages", effectiveOptions => Extract(ResolveSelector(selector, effectiveOptions), effectiveOptions), options);
    }

    /// <summary>Creates a new PDF with pages resolved from a document-relative selector deleted.</summary>
    public PdfDocument Delete(OfficePageSelector selector) {
        return Delete(ResolveSelector(selector, _document.ReadOptions));
    }

    /// <summary>Attempts to delete pages resolved from a document-relative selector.</summary>
    internal PdfOperationResult<PdfDocument> DeleteResult(OfficePageSelector selector, PdfLoadOptions? options = null) {
        Guard.NotNull(selector, nameof(selector));
        PdfLoadOptions? effectiveOptions = options ?? _document.ReadOptions;
        return _document.TryMutationOperation(
            "Delete pages",
            PdfPreflightCapability.ManipulatePages,
            PdfMutationOperation.OrganizePages,
            () => Delete(ResolveSelector(selector, effectiveOptions), effectiveOptions),
            options);
    }

    /// <summary>Creates a new PDF with pages copied in the order resolved by a document-relative selector.</summary>
    public PdfDocument Reorder(OfficePageSelector selector) {
        return Reorder(ResolveSelector(selector, _document.ReadOptions));
    }

    /// <summary>Attempts to reorder pages resolved from a document-relative selector.</summary>
    internal PdfOperationResult<PdfDocument> ReorderResult(OfficePageSelector selector, PdfLoadOptions? options = null) {
        Guard.NotNull(selector, nameof(selector));
        PdfLoadOptions? effectiveOptions = options ?? _document.ReadOptions;
        return _document.TryMutationOperation(
            "Reorder pages",
            PdfPreflightCapability.ManipulatePages,
            PdfMutationOperation.OrganizePages,
            () => Reorder(ResolveSelector(selector, effectiveOptions), effectiveOptions),
            options);
    }

    /// <summary>Creates a new PDF with pages resolved from a document-relative selector duplicated.</summary>
    public PdfDocument Duplicate(OfficePageSelector selector) {
        return Duplicate(ResolveSelector(selector, _document.ReadOptions));
    }

    /// <summary>Attempts to duplicate pages resolved from a document-relative selector.</summary>
    internal PdfOperationResult<PdfDocument> DuplicateResult(OfficePageSelector selector, PdfLoadOptions? options = null) {
        Guard.NotNull(selector, nameof(selector));
        PdfLoadOptions? effectiveOptions = options ?? _document.ReadOptions;
        return _document.TryMutationOperation(
            "Duplicate pages",
            PdfPreflightCapability.ManipulatePages,
            PdfMutationOperation.OrganizePages,
            () => Duplicate(ResolveSelector(selector, effectiveOptions), effectiveOptions),
            options);
    }

    /// <summary>Moves pages resolved from a document-relative selector before a one-based page number.</summary>
    public PdfDocument Move(int insertBeforePageNumber, OfficePageSelector selector) {
        return Move(insertBeforePageNumber, ResolveSelector(selector, _document.ReadOptions));
    }

    /// <summary>Attempts to move pages resolved from a document-relative selector.</summary>
    internal PdfOperationResult<PdfDocument> MoveResult(int insertBeforePageNumber, OfficePageSelector selector, PdfLoadOptions? options = null) {
        Guard.NotNull(selector, nameof(selector));
        PdfLoadOptions? effectiveOptions = options ?? _document.ReadOptions;
        return _document.TryMutationOperation(
            "Move pages",
            PdfPreflightCapability.ManipulatePages,
            PdfMutationOperation.OrganizePages,
            () => Move(insertBeforePageNumber, ResolveSelector(selector, effectiveOptions), effectiveOptions),
            options);
    }

    /// <summary>Rotates pages resolved from a document-relative selector.</summary>
    public PdfDocument Rotate(int rotationDegrees, OfficePageSelector selector) {
        return Rotate(rotationDegrees, ResolveSelector(selector, _document.ReadOptions));
    }

    /// <summary>Attempts to rotate pages resolved from a document-relative selector.</summary>
    internal PdfOperationResult<PdfDocument> RotateResult(int rotationDegrees, OfficePageSelector selector, PdfLoadOptions? options = null) {
        Guard.NotNull(selector, nameof(selector));
        PdfLoadOptions? effectiveOptions = options ?? _document.ReadOptions;
        return _document.TryMutationOperation(
            "Rotate pages",
            PdfPreflightCapability.ManipulatePages,
            PdfMutationOperation.RotatePages,
            () => Rotate(rotationDegrees, ResolveSelector(selector, effectiveOptions), effectiveOptions),
            options);
    }

    private PdfPageSelection ResolveSelector(OfficePageSelector selector, PdfLoadOptions? options, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        Guard.NotNull(selector, nameof(selector));
        int pageCount = GetSelectionPageCount(options, cancellationToken);
        if (pageCount < 1) {
            throw new InvalidOperationException("PDF does not contain any readable pages.");
        }

        return selector.ResolveSelection(pageCount);
    }

    // Selection resolution needs page geometry, not permission to extract logical content.
    private int GetSelectionPageCount(PdfLoadOptions? options, CancellationToken cancellationToken = default) =>
        (_document.GetOpenedReadDocumentFactory(options, cancellationToken)?.Invoke() ??
            PdfReadDocument.Open(_document.GetBytesForOperation(cancellationToken), options, cancellationToken)).Pages.Count;
}
