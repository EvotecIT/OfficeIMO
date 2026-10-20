using System.Threading;

namespace OfficeIMO.Pdf;

/// <summary>
/// Fluent page extraction and editing operations for a <see cref="PdfDocument"/>.
/// </summary>
public sealed partial class PdfDocumentPages {
    private readonly PdfDocument _document;

    internal PdfDocumentPages(PdfDocument document) {
        _document = document;
    }

    /// <summary>Places selected source pages onto printable sheets as vector page content and returns their cell mapping.</summary>
    public PdfImpositionResult ImposeNUp(PdfNUpOptions options, PdfPageSelection? selection = null) =>
        PdfPageImposer.ImposeNUp(_document.GetBytesForOperation(), options, selection, _document.ReadOptions);

    /// <summary>Creates a two-up duplex booklet in front/back print order. Source pages become visual sheet content.</summary>
    public PdfImpositionResult ImposeBooklet(PdfBookletOptions options, PdfPageSelection? selection = null) =>
        PdfPageImposer.ImposeBooklet(_document.GetBytesForOperation(), options, selection, _document.ReadOptions);

    /// <summary>
    /// Returns outline/bookmark-derived page ranges in document order.
    /// </summary>
    public IReadOnlyList<PdfBookmarkPageRange> BookmarkPageRanges(params string[] bookmarkTitles) {
        return BookmarkPageRanges(_document.ReadOptions, default, bookmarkTitles);
    }

    private IReadOnlyList<PdfBookmarkPageRange> BookmarkPageRanges(PdfLoadOptions? options, CancellationToken cancellationToken, params string[] bookmarkTitles) {
        PdfDocumentInfo info = _document.Inspect(options, cancellationToken);
        PdfBookmarkPageRange[] allRanges = BuildBookmarkPageRanges(info, cancellationToken);
        if (bookmarkTitles is null || bookmarkTitles.Length == 0) {
            return allRanges;
        }

        var selected = new List<PdfBookmarkPageRange>(bookmarkTitles.Length);
        for (int i = 0; i < bookmarkTitles.Length; i++) {
            cancellationToken.ThrowIfCancellationRequested();
            string title = bookmarkTitles[i];
            Guard.NotNullOrWhiteSpace(title, nameof(bookmarkTitles));
            PdfBookmarkPageRange? match = allRanges.FirstOrDefault(range => string.Equals(range.Title, title, StringComparison.Ordinal));
            if (match is null) {
                throw new ArgumentException("Bookmark title '" + title + "' was not found or does not resolve to a page.", nameof(bookmarkTitles));
            }

            selected.Add(match);
        }

        return selected.AsReadOnly();
    }

    /// <summary>
    /// Creates one PDF for each outline/bookmark-derived page range.
    /// </summary>
    public IReadOnlyList<PdfDocument> SplitByBookmarks(params string[] bookmarkTitles) {
        IReadOnlyList<PdfBookmarkPageRange> ranges = BookmarkPageRanges(bookmarkTitles);
        if (ranges.Count == 0) {
            throw new InvalidOperationException("PDF does not contain any readable bookmarks with page destinations.");
        }

        return Split(ranges.Select(range => range.PageRange));
    }

    private PdfDocument[] SplitByBookmarks(PdfLoadOptions? options, CancellationToken cancellationToken, params string[] bookmarkTitles) {
        IReadOnlyList<PdfBookmarkPageRange> ranges = BookmarkPageRanges(options, cancellationToken, bookmarkTitles);
        if (ranges.Count == 0) {
            throw new InvalidOperationException("PDF does not contain any readable bookmarks with page destinations.");
        }

        return Split(ranges.Select(range => range.PageRange), options, cancellationToken);
    }

    /// <summary>
    /// Attempts to create one PDF per page, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<PdfDocument>> SplitResult(PdfLoadOptions? options = null) {
        return TryPageExtractionOperation<IReadOnlyList<PdfDocument>>("Split pages", effectiveOptions => Split(effectiveOptions), options);
    }

    /// <summary>
    /// Attempts to create PDFs containing consecutive groups of the requested page count.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<PdfDocument>> SplitResult(int pagesPerDocument, PdfLoadOptions? options = null) {
        return TryPageExtractionOperation<IReadOnlyList<PdfDocument>>("Split page groups", effectiveOptions => Split(pagesPerDocument, effectiveOptions), options);
    }

    /// <summary>
    /// Attempts to create one PDF for each supplied page selection.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<PdfDocument>> SplitResult(IReadOnlyList<PdfPageSelection> selections, PdfLoadOptions? options = null) {
        Guard.NotNull(selections, nameof(selections));
        if (selections.Count == 0) {
            return TryPageExtractionOperation<IReadOnlyList<PdfDocument>>(
                "Split page selections",
                _ => throw new ArgumentException("At least one page selection must be specified.", nameof(selections)),
                options);
        }

        return TryPageExtractionOperation<IReadOnlyList<PdfDocument>>("Split page selections", effectiveOptions => Split(selections.ToArray(), effectiveOptions), options);
    }

    /// <summary>
    /// Attempts to create one PDF for each supplied inclusive page range, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<PdfDocument>> SplitResult(IEnumerable<PdfPageRange> pageRanges, PdfLoadOptions? options = null) {
        Guard.NotNull(pageRanges, nameof(pageRanges));
        return TryPageExtractionOperation<IReadOnlyList<PdfDocument>>("Split page ranges", effectiveOptions => Split(pageRanges, effectiveOptions), options);
    }

    /// <summary>
    /// Attempts to create one PDF for each parsed page range, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<PdfDocument>> SplitResult(string pageRanges, PdfLoadOptions? options = null) {
        return TryPageExtractionOperation<IReadOnlyList<PdfDocument>>("Split page ranges", effectiveOptions => Split(PdfPageRange.ParseMany(pageRanges), effectiveOptions), options);
    }

    /// <summary>
    /// Attempts to create one PDF for each outline/bookmark-derived page range.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<PdfDocument>> SplitByBookmarksResult(IReadOnlyList<string>? bookmarkTitles = null, PdfLoadOptions? options = null) {
        return TryPageExtractionOperation<IReadOnlyList<PdfDocument>>(
            "Split bookmarks",
            effectiveOptions => SplitByBookmarks(effectiveOptions, default, bookmarkTitles is null ? Array.Empty<string>() : bookmarkTitles.ToArray()),
            options);
    }

    /// <summary>
    /// Creates a new PDF with selected pages deleted.
    /// </summary>
    public PdfDocument Delete(params int[] pageNumbers) {
        return _document.ApplyMutation(input => PdfPageEditor.DeletePagesWithReadOptions(input, _document.ReadOptions, pageNumbers));
    }

    /// <summary>
    /// Creates a new PDF with one inclusive page range deleted.
    /// </summary>
    public PdfDocument Delete(PdfPageRange pageRange) {
        return _document.ApplyMutation(input => PdfPageEditor.DeletePagesWithReadOptions(input, _document.ReadOptions, pageRange.ToPageNumbers()));
    }

    /// <summary>
    /// Creates a new PDF with selected pages deleted.
    /// </summary>
    public PdfDocument Delete(PdfPageSelection selection) {
        Guard.NotNull(selection, nameof(selection));
        return Delete(selection, _document.ReadOptions);
    }

    private PdfDocument Delete(PdfPageSelection selection, PdfLoadOptions? readOptions, CancellationToken cancellationToken = default) =>
        _document.ApplyMutation(input => PdfPageEditor.DeletePagesWithReadOptions(input, readOptions, cancellationToken, selection.ToPageNumbers(GetSelectionPageCount(readOptions, cancellationToken), nameof(selection))), readOptions, cancellationToken: cancellationToken);

    /// <summary>
    /// Attempts to create a new PDF with selected pages deleted, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> DeleteResult(PdfPageSelection selection, PdfLoadOptions? options = null) {
        Guard.NotNull(selection, nameof(selection));
        return _document.TryMutationOperation("Delete pages", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.OrganizePages, () => Delete(selection, options ?? _document.ReadOptions), options);
    }

    /// <summary>
    /// Creates a new PDF with comma- or semicolon-separated inclusive page ranges deleted.
    /// </summary>
    public PdfDocument Delete(string pageRanges) {
        return Delete(PdfPageSelection.Parse(pageRanges));
    }

    /// <summary>
    /// Attempts to create a new PDF with pages described by page ranges deleted, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> DeleteResult(string pageRanges, PdfLoadOptions? options = null) {
        return _document.TryMutationOperation("Delete pages", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.OrganizePages, () => Delete(PdfPageSelection.Parse(pageRanges), options ?? _document.ReadOptions), options);
    }

    /// <summary>
    /// Creates a new PDF with every page copied in the specified one-based order.
    /// </summary>
    public PdfDocument Reorder(params int[] pageNumbers) {
        return _document.ApplyMutation(input => PdfPageEditor.ReorderPagesWithReadOptions(input, _document.ReadOptions, pageNumbers));
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
        return Reorder(selection, _document.ReadOptions);
    }

    private PdfDocument Reorder(PdfPageSelection selection, PdfLoadOptions? readOptions, CancellationToken cancellationToken = default) =>
        _document.ApplyMutation(input => PdfPageEditor.ReorderPagesWithReadOptions(input, readOptions, cancellationToken, selection.ToPageNumbers(GetSelectionPageCount(readOptions, cancellationToken), nameof(selection))), readOptions, cancellationToken: cancellationToken);

    /// <summary>
    /// Attempts to create a new PDF with every page copied in the selected one-based order, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> ReorderResult(PdfPageSelection selection, PdfLoadOptions? options = null) {
        Guard.NotNull(selection, nameof(selection));
        return _document.TryMutationOperation("Reorder pages", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.OrganizePages, () => Reorder(selection, options ?? _document.ReadOptions), options);
    }

    /// <summary>
    /// Attempts to create a new PDF with pages copied in parsed page-range order, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> ReorderResult(string pageRanges, PdfLoadOptions? options = null) {
        return _document.TryMutationOperation("Reorder pages", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.OrganizePages, () => Reorder(PdfPageSelection.Parse(pageRanges), options ?? _document.ReadOptions), options);
    }

    /// <summary>
    /// Creates a new PDF with selected pages duplicated immediately after each source page.
    /// </summary>
    public PdfDocument Duplicate(params int[] pageNumbers) {
        return _document.ApplyMutation(input => PdfPageEditor.DuplicatePagesWithReadOptions(input, _document.ReadOptions, pageNumbers));
    }

    /// <summary>
    /// Creates a new PDF with one inclusive page range duplicated.
    /// </summary>
    public PdfDocument Duplicate(PdfPageRange pageRange) {
        return _document.ApplyMutation(input => PdfPageEditor.DuplicatePagesWithReadOptions(input, _document.ReadOptions, pageRange.ToPageNumbers()));
    }

    /// <summary>
    /// Creates a new PDF with selected pages duplicated immediately after each source page.
    /// </summary>
    public PdfDocument Duplicate(PdfPageSelection selection) {
        Guard.NotNull(selection, nameof(selection));
        return Duplicate(selection, _document.ReadOptions);
    }

    private PdfDocument Duplicate(PdfPageSelection selection, PdfLoadOptions? readOptions, CancellationToken cancellationToken = default) =>
        _document.ApplyMutation(input => PdfPageEditor.DuplicatePagesWithReadOptions(input, readOptions, cancellationToken, selection.ToPageNumbers(GetSelectionPageCount(readOptions, cancellationToken), nameof(selection))), readOptions, cancellationToken: cancellationToken);

    /// <summary>
    /// Attempts to create a new PDF with selected pages duplicated, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> DuplicateResult(PdfPageSelection selection, PdfLoadOptions? options = null) {
        Guard.NotNull(selection, nameof(selection));
        return _document.TryMutationOperation("Duplicate pages", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.OrganizePages, () => Duplicate(selection, options ?? _document.ReadOptions), options);
    }

    /// <summary>
    /// Creates a new PDF with parsed page ranges duplicated.
    /// </summary>
    public PdfDocument Duplicate(string pageRanges) {
        return Duplicate(PdfPageSelection.Parse(pageRanges));
    }

    /// <summary>
    /// Attempts to create a new PDF with parsed page ranges duplicated, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> DuplicateResult(string pageRanges, PdfLoadOptions? options = null) {
        return _document.TryMutationOperation("Duplicate pages", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.OrganizePages, () => Duplicate(PdfPageSelection.Parse(pageRanges), options ?? _document.ReadOptions), options);
    }

    /// <summary>
    /// Creates a new PDF with selected pages moved before the supplied one-based page number.
    /// Use page count + 1 to move pages to the end.
    /// </summary>
    public PdfDocument Move(int insertBeforePageNumber, params int[] pageNumbers) {
        return _document.ApplyMutation(input => PdfPageEditor.MovePagesWithReadOptions(input, insertBeforePageNumber, _document.ReadOptions, pageNumbers));
    }

    /// <summary>
    /// Creates a new PDF with one inclusive page range moved before the supplied one-based page number.
    /// Use page count + 1 to move pages to the end.
    /// </summary>
    public PdfDocument Move(int insertBeforePageNumber, PdfPageRange pageRange) {
        return _document.ApplyMutation(input => PdfPageEditor.MovePagesWithReadOptions(input, insertBeforePageNumber, _document.ReadOptions, pageRange.ToPageNumbers()));
    }

    /// <summary>
    /// Creates a new PDF with selected pages moved before the supplied one-based page number.
    /// Use page count + 1 to move pages to the end.
    /// </summary>
    public PdfDocument Move(int insertBeforePageNumber, PdfPageSelection selection) {
        Guard.NotNull(selection, nameof(selection));
        return Move(insertBeforePageNumber, selection, _document.ReadOptions);
    }

    private PdfDocument Move(int insertBeforePageNumber, PdfPageSelection selection, PdfLoadOptions? readOptions, CancellationToken cancellationToken = default) =>
        _document.ApplyMutation(input => PdfPageEditor.MovePagesWithReadOptions(input, insertBeforePageNumber, readOptions, cancellationToken, selection.ToPageNumbers(GetSelectionPageCount(readOptions, cancellationToken), nameof(selection))), readOptions, cancellationToken: cancellationToken);

    /// <summary>
    /// Attempts to create a new PDF with selected pages moved before the supplied one-based page number, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> MoveResult(int insertBeforePageNumber, PdfPageSelection selection, PdfLoadOptions? options = null) {
        Guard.NotNull(selection, nameof(selection));
        return _document.TryMutationOperation("Move pages", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.OrganizePages, () => Move(insertBeforePageNumber, selection, options ?? _document.ReadOptions), options);
    }

    /// <summary>
    /// Creates a new PDF with parsed page ranges moved before the supplied one-based page number.
    /// Use page count + 1 to move pages to the end.
    /// </summary>
    public PdfDocument Move(int insertBeforePageNumber, string pageRanges) {
        return Move(insertBeforePageNumber, PdfPageSelection.Parse(pageRanges));
    }

    /// <summary>
    /// Attempts to create a new PDF with parsed page ranges moved before the supplied one-based page number, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> MoveResult(int insertBeforePageNumber, string pageRanges, PdfLoadOptions? options = null) {
        return _document.TryMutationOperation("Move pages", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.OrganizePages, () => Move(insertBeforePageNumber, PdfPageSelection.Parse(pageRanges), options ?? _document.ReadOptions), options);
    }

    /// <summary>
    /// Creates a new PDF with selected pages rotated. Supplying no page numbers rotates every page.
    /// </summary>
    public PdfDocument Rotate(int rotationDegrees, params int[] pageNumbers) {
        return _document.ApplyMutation(input => PdfPageEditor.RotatePagesWithReadOptions(input, rotationDegrees, _document.ReadOptions, pageNumbers));
    }

    /// <summary>
    /// Creates a new PDF with one inclusive page range rotated.
    /// </summary>
    public PdfDocument Rotate(int rotationDegrees, PdfPageRange pageRange) {
        return _document.ApplyMutation(input => PdfPageEditor.RotatePagesWithReadOptions(input, rotationDegrees, _document.ReadOptions, pageRange.ToPageNumbers()));
    }

    /// <summary>
    /// Creates a new PDF with selected pages rotated.
    /// </summary>
    public PdfDocument Rotate(int rotationDegrees, PdfPageSelection selection) {
        Guard.NotNull(selection, nameof(selection));
        return Rotate(rotationDegrees, selection, _document.ReadOptions);
    }

    private PdfDocument Rotate(int rotationDegrees, PdfPageSelection selection, PdfLoadOptions? readOptions, CancellationToken cancellationToken = default) =>
        _document.ApplyMutation(input => PdfPageEditor.RotatePagesWithReadOptions(input, rotationDegrees, readOptions, cancellationToken, selection.ToPageNumbers(GetSelectionPageCount(readOptions, cancellationToken), nameof(selection))), readOptions, cancellationToken: cancellationToken);

    /// <summary>
    /// Attempts to create a new PDF with selected pages rotated, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> RotateResult(int rotationDegrees, PdfPageSelection selection, PdfLoadOptions? options = null) {
        Guard.NotNull(selection, nameof(selection));
        return _document.TryMutationOperation("Rotate pages", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.RotatePages, () => Rotate(rotationDegrees, selection, options ?? _document.ReadOptions), options);
    }

    /// <summary>
    /// Creates a new PDF with parsed page ranges rotated.
    /// </summary>
    public PdfDocument Rotate(int rotationDegrees, string pageRanges) {
        return Rotate(rotationDegrees, PdfPageSelection.Parse(pageRanges));
    }

    /// <summary>
    /// Attempts to create a new PDF with parsed page ranges rotated, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> RotateResult(int rotationDegrees, string pageRanges, PdfLoadOptions? options = null) {
        return _document.TryMutationOperation("Rotate pages", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.RotatePages, () => Rotate(rotationDegrees, PdfPageSelection.Parse(pageRanges), options ?? _document.ReadOptions), options);
    }

    /// <summary>Sets a typed boundary box for selected pages, or every page when no page numbers are supplied.</summary>
    public PdfDocument SetPageBox(PdfPageBoundaryBox box, double left, double bottom, double right, double top, params int[] pageNumbers) {
        return _document.ApplyMutation(input => PdfPageEditor.SetPageBoxWithReadOptions(input, box, left, bottom, right, top, _document.ReadOptions, pageNumbers));
    }

    /// <summary>Resizes selected pages to the supplied page size.</summary>
    public PdfDocument Resize(PageSize pageSize, params int[] pageNumbers) {
        return _document.ApplyMutation(input => PdfPageEditor.ResizePages(input, pageSize, _document.ReadOptions, pageNumbers));
    }

    /// <summary>Resizes selected pages with explicit fit, fill, alignment, and scaling settings.</summary>
    public PdfDocument Resize(PdfPageResizeOptions options, params int[] pageNumbers) {
        Guard.NotNull(options, nameof(options));
        return _document.ApplyMutation(input => PdfPageEditor.ResizePages(input, options, _document.ReadOptions, pageNumbers), cancellationToken: options.CancellationToken);
    }

    /// <summary>Sets the media box for selected pages.</summary>
    public PdfDocument SetMediaBox(double left, double bottom, double right, double top, params int[] pageNumbers) =>
        SetPageBox(PdfPageBoundaryBox.MediaBox, left, bottom, right, top, pageNumbers);

    /// <summary>Sets the crop box for selected pages.</summary>
    public PdfDocument SetCropBox(double left, double bottom, double right, double top, params int[] pageNumbers) =>
        SetPageBox(PdfPageBoundaryBox.CropBox, left, bottom, right, top, pageNumbers);

    /// <summary>
    /// Non-destructively crops selected pages and translates the chosen source rectangle to a zero-based page origin.
    /// Content outside the rectangle remains in source streams but is clipped from display.
    /// </summary>
    public PdfDocument CropAndTranslate(double left, double bottom, double right, double top, params int[] pageNumbers) =>
        _document.ApplyMutation(input => PdfPageEditor.CropAndTranslatePages(
            input, left, bottom, right, top, _document.ReadOptions, default, pageNumbers));

    /// <summary>Destructively crops selected pages by replacing the retained visual rectangle with a validated opaque raster page.</summary>
    public PdfDestructiveCropResult DestructiveCrop(double left, double bottom, double right, double top, PdfDestructiveCropOptions? options = null, params int[] pageNumbers) =>
        PdfPageEditor.DestructiveCropPages(_document.GetBytesForOperation(options?.CancellationToken ?? default), left, bottom, right, top, options, _document.ReadOptions, pageNumbers);

    /// <summary>Sets the bleed box for selected pages.</summary>
    public PdfDocument SetBleedBox(double left, double bottom, double right, double top, params int[] pageNumbers) =>
        SetPageBox(PdfPageBoundaryBox.BleedBox, left, bottom, right, top, pageNumbers);

    /// <summary>Sets the trim box for selected pages.</summary>
    public PdfDocument SetTrimBox(double left, double bottom, double right, double top, params int[] pageNumbers) =>
        SetPageBox(PdfPageBoundaryBox.TrimBox, left, bottom, right, top, pageNumbers);

    /// <summary>Sets the art box for selected pages.</summary>
    public PdfDocument SetArtBox(double left, double bottom, double right, double top, params int[] pageNumbers) =>
        SetPageBox(PdfPageBoundaryBox.ArtBox, left, bottom, right, top, pageNumbers);

    private static PdfBookmarkPageRange[] BuildBookmarkPageRanges(PdfDocumentInfo info, CancellationToken cancellationToken) {
        var outlines = new List<PdfOutlineItem>();
        FlattenOutlines(info.Outlines, outlines, cancellationToken);

        var anchors = new List<PdfOutlineItem>();
        for (int i = 0; i < outlines.Count; i++) {
            cancellationToken.ThrowIfCancellationRequested();
            if (outlines[i].PageNumber.HasValue) {
                int pageNumber = outlines[i].PageNumber!.Value;
                if (pageNumber >= 1 && pageNumber <= info.PageCount) {
                    anchors.Add(outlines[i]);
                }
            }
        }

        if (anchors.Count == 0) {
            return Array.Empty<PdfBookmarkPageRange>();
        }

        anchors = anchors
            .OrderBy(anchor => anchor.PageNumber!.Value)
            .ToList();

        var ranges = new List<PdfBookmarkPageRange>(anchors.Count);
        for (int i = 0; i < anchors.Count; i++) {
            cancellationToken.ThrowIfCancellationRequested();
            PdfOutlineItem anchor = anchors[i];
            int firstPage = anchor.PageNumber!.Value;
            int lastPage = info.PageCount;
            for (int j = i + 1; j < anchors.Count; j++) {
                int nextPage = anchors[j].PageNumber!.Value;
                if (nextPage > firstPage) {
                    lastPage = nextPage - 1;
                    break;
                }
            }

            if (lastPage < firstPage) {
                lastPage = firstPage;
            }

            ranges.Add(new PdfBookmarkPageRange(anchor.Title, anchor.Level, PdfPageRange.From(firstPage, lastPage), anchor));
        }

        return ranges.ToArray();
    }

    private static void FlattenOutlines(IReadOnlyList<PdfOutlineItem> source, List<PdfOutlineItem> destination, CancellationToken cancellationToken) {
        for (int i = 0; i < source.Count; i++) {
            cancellationToken.ThrowIfCancellationRequested();
            PdfOutlineItem outline = source[i];
            destination.Add(outline);
            if (outline.Children.Count > 0) {
                FlattenOutlines(outline.Children, destination, cancellationToken);
            }
        }
    }

    private PdfOperationResult<T> TryPageExtractionOperation<T>(
        string operationName,
        Func<PdfLoadOptions?, T> operation,
        PdfLoadOptions? options, CancellationToken cancellationToken = default) where T : class {
        PdfLoadOptions? effectiveOptions = options ?? _document.ReadOptions;
        PdfMutationPlan plan = _document.PlanMutation(PdfMutationOperation.ExtractPages, fieldNames: null, options: effectiveOptions, executionPreference: PdfMutationExecutionPreference.Automatic, cancellationToken: cancellationToken);
        if (!plan.CanExecute) {
            return PdfOperationResult<T>.MutationBlocked(operationName, PdfPreflightCapability.ManipulatePages, plan);
        }

        try {
            return PdfOperationResult<T>.MutationSuccess(
                operationName,
                PdfPreflightCapability.ManipulatePages,
                plan,
                operation(effectiveOptions));
        } catch (OperationCanceledException) {
            throw;
        } catch (Exception ex) {
            return PdfOperationResult<T>.MutationFailed(
                operationName,
                PdfPreflightCapability.ManipulatePages,
                plan,
                ex);
        }
    }
}
