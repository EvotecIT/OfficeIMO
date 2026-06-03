namespace OfficeIMO.Pdf;

/// <summary>
/// Fluent page extraction and editing operations for a <see cref="PdfDocument"/>.
/// </summary>
public sealed class PdfDocumentPages {
    private readonly PdfDocument _document;

    internal PdfDocumentPages(PdfDocument document) {
        _document = document;
    }

    /// <summary>
    /// Creates a new PDF containing selected pages in caller order.
    /// </summary>
    public PdfDocument Extract(params int[] pageNumbers) {
        return PdfDocument.FromBytes(PdfPageExtractor.ExtractPages(_document.Snapshot(), pageNumbers));
    }

    /// <summary>
    /// Creates a new PDF containing one inclusive one-based page range.
    /// </summary>
    public PdfDocument Extract(PdfPageRange pageRange) {
        return PdfDocument.FromBytes(PdfPageExtractor.ExtractPageRange(_document.Snapshot(), pageRange));
    }

    /// <summary>
    /// Creates a new PDF containing selected pages in caller order.
    /// </summary>
    public PdfDocument Extract(PdfPageSelection selection) {
        Guard.NotNull(selection, nameof(selection));
        return PdfDocument.FromBytes(PdfPageExtractor.ExtractPageRanges(_document.Snapshot(), selection.ToRanges()));
    }

    /// <summary>
    /// Attempts to create a new PDF containing selected pages in caller order, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryExtract(PdfPageSelection selection, PdfReadOptions? options = null) {
        Guard.NotNull(selection, nameof(selection));
        return _document.TryOperation("Extract pages", PdfPreflightCapability.ManipulatePages, () => Extract(selection), options);
    }

    /// <summary>
    /// Creates a new PDF containing comma- or semicolon-separated inclusive page ranges.
    /// </summary>
    public PdfDocument Extract(string pageRanges) {
        return Extract(PdfPageSelection.Parse(pageRanges));
    }

    /// <summary>
    /// Creates one PDF per page.
    /// </summary>
    public IReadOnlyList<PdfDocument> Split() {
        return PdfPageExtractor.SplitPages(_document.Snapshot())
            .Select(PdfDocument.FromBytes)
            .ToArray();
    }

    /// <summary>
    /// Attempts to create one PDF per page, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<PdfDocument>> TrySplit(PdfReadOptions? options = null) {
        return _document.TryOperation("Split pages", PdfPreflightCapability.ManipulatePages, Split, options);
    }

    /// <summary>
    /// Creates a new PDF with selected pages deleted.
    /// </summary>
    public PdfDocument Delete(params int[] pageNumbers) {
        return PdfDocument.FromBytes(PdfPageEditor.DeletePages(_document.Snapshot(), pageNumbers));
    }

    /// <summary>
    /// Creates a new PDF with one inclusive page range deleted.
    /// </summary>
    public PdfDocument Delete(PdfPageRange pageRange) {
        return PdfDocument.FromBytes(PdfPageEditor.DeletePageRange(_document.Snapshot(), pageRange));
    }

    /// <summary>
    /// Creates a new PDF with selected pages deleted.
    /// </summary>
    public PdfDocument Delete(PdfPageSelection selection) {
        Guard.NotNull(selection, nameof(selection));
        return PdfDocument.FromBytes(PdfPageEditor.DeletePageRanges(_document.Snapshot(), selection.ToRanges()));
    }

    /// <summary>
    /// Attempts to create a new PDF with selected pages deleted, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryDelete(PdfPageSelection selection, PdfReadOptions? options = null) {
        Guard.NotNull(selection, nameof(selection));
        return _document.TryOperation("Delete pages", PdfPreflightCapability.ManipulatePages, () => Delete(selection), options);
    }

    /// <summary>
    /// Creates a new PDF with comma- or semicolon-separated inclusive page ranges deleted.
    /// </summary>
    public PdfDocument Delete(string pageRanges) {
        return Delete(PdfPageSelection.Parse(pageRanges));
    }

    /// <summary>
    /// Creates a new PDF with every page copied in the specified one-based order.
    /// </summary>
    public PdfDocument Reorder(params int[] pageNumbers) {
        return PdfDocument.FromBytes(PdfPageEditor.ReorderPages(_document.Snapshot(), pageNumbers));
    }

    /// <summary>
    /// Creates a new PDF with every page copied in parsed page-range order.
    /// </summary>
    public PdfDocument Reorder(string pageRanges) {
        return Reorder(PdfPageSelection.Parse(pageRanges));
    }

    /// <summary>
    /// Creates a new PDF with every page copied in the selected one-based order.
    /// </summary>
    public PdfDocument Reorder(PdfPageSelection selection) {
        Guard.NotNull(selection, nameof(selection));
        return PdfDocument.FromBytes(PdfPageEditor.ReorderPageRanges(_document.Snapshot(), selection.ToRanges()));
    }

    /// <summary>
    /// Attempts to create a new PDF with every page copied in the selected one-based order, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryReorder(PdfPageSelection selection, PdfReadOptions? options = null) {
        Guard.NotNull(selection, nameof(selection));
        return _document.TryOperation("Reorder pages", PdfPreflightCapability.ManipulatePages, () => Reorder(selection), options);
    }

    /// <summary>
    /// Creates a new PDF with selected pages duplicated immediately after each source page.
    /// </summary>
    public PdfDocument Duplicate(params int[] pageNumbers) {
        return PdfDocument.FromBytes(PdfPageEditor.DuplicatePages(_document.Snapshot(), pageNumbers));
    }

    /// <summary>
    /// Creates a new PDF with one inclusive page range duplicated.
    /// </summary>
    public PdfDocument Duplicate(PdfPageRange pageRange) {
        return PdfDocument.FromBytes(PdfPageEditor.DuplicatePageRange(_document.Snapshot(), pageRange));
    }

    /// <summary>
    /// Creates a new PDF with selected pages duplicated immediately after each source page.
    /// </summary>
    public PdfDocument Duplicate(PdfPageSelection selection) {
        Guard.NotNull(selection, nameof(selection));
        return PdfDocument.FromBytes(PdfPageEditor.DuplicatePageRanges(_document.Snapshot(), selection.ToRanges()));
    }

    /// <summary>
    /// Attempts to create a new PDF with selected pages duplicated, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryDuplicate(PdfPageSelection selection, PdfReadOptions? options = null) {
        Guard.NotNull(selection, nameof(selection));
        return _document.TryOperation("Duplicate pages", PdfPreflightCapability.ManipulatePages, () => Duplicate(selection), options);
    }

    /// <summary>
    /// Creates a new PDF with parsed page ranges duplicated.
    /// </summary>
    public PdfDocument Duplicate(string pageRanges) {
        return Duplicate(PdfPageSelection.Parse(pageRanges));
    }

    /// <summary>
    /// Creates a new PDF with selected pages moved before the supplied one-based page number.
    /// Use page count + 1 to move pages to the end.
    /// </summary>
    public PdfDocument Move(int insertBeforePageNumber, params int[] pageNumbers) {
        return PdfDocument.FromBytes(PdfPageEditor.MovePages(_document.Snapshot(), insertBeforePageNumber, pageNumbers));
    }

    /// <summary>
    /// Creates a new PDF with one inclusive page range moved before the supplied one-based page number.
    /// Use page count + 1 to move pages to the end.
    /// </summary>
    public PdfDocument Move(int insertBeforePageNumber, PdfPageRange pageRange) {
        return PdfDocument.FromBytes(PdfPageEditor.MovePageRange(_document.Snapshot(), insertBeforePageNumber, pageRange));
    }

    /// <summary>
    /// Creates a new PDF with selected pages moved before the supplied one-based page number.
    /// Use page count + 1 to move pages to the end.
    /// </summary>
    public PdfDocument Move(int insertBeforePageNumber, PdfPageSelection selection) {
        Guard.NotNull(selection, nameof(selection));
        return PdfDocument.FromBytes(PdfPageEditor.MovePageRanges(_document.Snapshot(), insertBeforePageNumber, selection.ToRanges()));
    }

    /// <summary>
    /// Attempts to create a new PDF with selected pages moved before the supplied one-based page number, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryMove(int insertBeforePageNumber, PdfPageSelection selection, PdfReadOptions? options = null) {
        Guard.NotNull(selection, nameof(selection));
        return _document.TryOperation("Move pages", PdfPreflightCapability.ManipulatePages, () => Move(insertBeforePageNumber, selection), options);
    }

    /// <summary>
    /// Creates a new PDF with parsed page ranges moved before the supplied one-based page number.
    /// Use page count + 1 to move pages to the end.
    /// </summary>
    public PdfDocument Move(int insertBeforePageNumber, string pageRanges) {
        return Move(insertBeforePageNumber, PdfPageSelection.Parse(pageRanges));
    }

    /// <summary>
    /// Creates a new PDF with selected pages rotated. Supplying no page numbers rotates every page.
    /// </summary>
    public PdfDocument Rotate(int rotationDegrees, params int[] pageNumbers) {
        return PdfDocument.FromBytes(PdfPageEditor.RotatePages(_document.Snapshot(), rotationDegrees, pageNumbers));
    }

    /// <summary>
    /// Creates a new PDF with one inclusive page range rotated.
    /// </summary>
    public PdfDocument Rotate(int rotationDegrees, PdfPageRange pageRange) {
        return PdfDocument.FromBytes(PdfPageEditor.RotatePageRange(_document.Snapshot(), rotationDegrees, pageRange));
    }

    /// <summary>
    /// Creates a new PDF with selected pages rotated.
    /// </summary>
    public PdfDocument Rotate(int rotationDegrees, PdfPageSelection selection) {
        Guard.NotNull(selection, nameof(selection));
        return PdfDocument.FromBytes(PdfPageEditor.RotatePageRanges(_document.Snapshot(), rotationDegrees, selection.ToRanges()));
    }

    /// <summary>
    /// Attempts to create a new PDF with selected pages rotated, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryRotate(int rotationDegrees, PdfPageSelection selection, PdfReadOptions? options = null) {
        Guard.NotNull(selection, nameof(selection));
        return _document.TryOperation("Rotate pages", PdfPreflightCapability.ManipulatePages, () => Rotate(rotationDegrees, selection), options);
    }

    /// <summary>
    /// Creates a new PDF with parsed page ranges rotated.
    /// </summary>
    public PdfDocument Rotate(int rotationDegrees, string pageRanges) {
        return Rotate(rotationDegrees, PdfPageSelection.Parse(pageRanges));
    }
}
