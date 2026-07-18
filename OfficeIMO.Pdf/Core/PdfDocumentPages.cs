namespace OfficeIMO.Pdf;

/// <summary>
/// Fluent page extraction and editing operations for a <see cref="PdfDocument"/>.
/// </summary>
public sealed partial class PdfDocumentPages {
    private readonly PdfDocument _document;

    internal PdfDocumentPages(PdfDocument document) {
        _document = document;
    }

    /// <summary>
    /// Creates a new PDF containing selected pages in caller order.
    /// </summary>
    public PdfDocument Extract(params int[] pageNumbers) {
        return _document.ApplyMutation(input => PdfPageExtractor.ExtractPages(input, _document.ReadOptions, pageNumbers));
    }

    /// <summary>
    /// Creates a new PDF containing one inclusive one-based page range.
    /// </summary>
    public PdfDocument Extract(PdfPageRange pageRange) {
        return _document.ApplyMutation(input => PdfPageExtractor.ExtractPages(input, pageRange.ToPageNumbers(), _document.ReadOptions));
    }

    /// <summary>
    /// Creates a new PDF containing selected pages in caller order.
    /// </summary>
    public PdfDocument Extract(PdfPageSelection selection) {
        Guard.NotNull(selection, nameof(selection));
        return Extract(selection, _document.ReadOptions);
    }

    private PdfDocument Extract(PdfPageSelection selection, PdfReadOptions? options) {
        Guard.NotNull(selection, nameof(selection));
        return _document.ApplyMutation(input => PdfPageExtractor.ExtractPageRanges(input, selection.ToRanges(), options));
    }

    /// <summary>
    /// Attempts to create a new PDF containing selected pages in caller order, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryExtract(PdfPageSelection selection, PdfReadOptions? options = null) {
        Guard.NotNull(selection, nameof(selection));
        return TryPageExtractionOperation("Extract pages", effectiveOptions => Extract(selection, effectiveOptions), options);
    }

    /// <summary>
    /// Creates a new PDF containing comma- or semicolon-separated inclusive page ranges.
    /// </summary>
    public PdfDocument Extract(string pageRanges) {
        return Extract(PdfPageSelection.Parse(pageRanges));
    }

    /// <summary>
    /// Attempts to create a new PDF containing pages described by page ranges, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryExtract(string pageRanges, PdfReadOptions? options = null) {
        return TryPageExtractionOperation("Extract pages", effectiveOptions => Extract(PdfPageSelection.Parse(pageRanges), effectiveOptions), options);
    }

    /// <summary>
    /// Creates one PDF per page.
    /// </summary>
    public IReadOnlyList<PdfDocument> Split() {
        return Split(_document.ReadOptions);
    }

    private PdfDocument[] Split(PdfReadOptions? options) {
        byte[] input = _document.GetBytesForOperation();
        return AdoptSplitOutputs(input, PdfPageExtractor.SplitPages(input, options), options);
    }

    /// <summary>
    /// Creates PDFs containing consecutive groups of the requested page count.
    /// </summary>
    public IReadOnlyList<PdfDocument> Split(int pagesPerDocument) {
        if (pagesPerDocument <= 0) {
            throw new ArgumentOutOfRangeException(nameof(pagesPerDocument), pagesPerDocument, "Pages per document must be greater than zero.");
        }

        return Split(pagesPerDocument, _document.ReadOptions);
    }

    private PdfDocument[] Split(int pagesPerDocument, PdfReadOptions? options) {
        if (pagesPerDocument <= 0) {
            throw new ArgumentOutOfRangeException(nameof(pagesPerDocument), pagesPerDocument, "Pages per document must be greater than zero.");
        }

        int pageCount = _document.Inspect(options).PageCount;
        if (pageCount == 0) {
            throw new InvalidOperationException("PDF does not contain any readable pages.");
        }

        var ranges = new List<PdfPageRange>();
        for (int firstPage = 1; firstPage <= pageCount; firstPage += pagesPerDocument) {
            int lastPage = Math.Min(firstPage + pagesPerDocument - 1, pageCount);
            ranges.Add(PdfPageRange.From(firstPage, lastPage));
        }

        return Split(ranges, options);
    }

    /// <summary>
    /// Creates one PDF for each supplied page selection.
    /// </summary>
    public IReadOnlyList<PdfDocument> Split(params PdfPageSelection[] selections) {
        Guard.NotNull(selections, nameof(selections));
        if (selections.Length == 0) {
            throw new ArgumentException("At least one page selection must be specified.", nameof(selections));
        }

        return Split(selections, _document.ReadOptions);
    }

    private PdfDocument[] Split(PdfPageSelection[] selections, PdfReadOptions? options) {
        byte[] input = _document.GetBytesForOperation();
        var outputs = new byte[selections.Length][];
        for (int i = 0; i < selections.Length; i++) {
            Guard.NotNull(selections[i], nameof(selections));
            outputs[i] = PdfPageExtractor.ExtractPageRanges(input, selections[i].ToRanges(), options);
        }

        return AdoptSplitOutputs(input, outputs, options);
    }

    /// <summary>
    /// Creates one PDF for each supplied inclusive page range.
    /// </summary>
    public IReadOnlyList<PdfDocument> Split(IEnumerable<PdfPageRange> pageRanges) {
        Guard.NotNull(pageRanges, nameof(pageRanges));
        PdfPageRange[] ranges = pageRanges.ToArray();
        if (ranges.Length == 0) {
            throw new ArgumentException("At least one page range must be specified.", nameof(pageRanges));
        }

        return Split(ranges, _document.ReadOptions);
    }

    private PdfDocument[] Split(IEnumerable<PdfPageRange> pageRanges, PdfReadOptions? options) {
        Guard.NotNull(pageRanges, nameof(pageRanges));
        PdfPageRange[] ranges = pageRanges.ToArray();
        if (ranges.Length == 0) {
            throw new ArgumentException("At least one page range must be specified.", nameof(pageRanges));
        }

        byte[] input = _document.GetBytesForOperation();
        return AdoptSplitOutputs(input, PdfPageExtractor.SplitPageRanges(input, ranges, options), options);
    }

    private PdfDocument[] AdoptSplitOutputs(
        byte[] input,
        IEnumerable<byte[]> outputs,
        PdfReadOptions? options) {
        PdfArtifactSnapshot inputArtifact = _document.Pipeline.Output ??
            PdfArtifactSnapshot.Capture(input, _document.ReadOptions);
        return outputs
            .Select(output => _document.WithBytes(input, inputArtifact, output, options, "Split"))
            .ToArray();
    }

    /// <summary>
    /// Creates one PDF for each comma- or semicolon-separated inclusive page range.
    /// </summary>
    public IReadOnlyList<PdfDocument> Split(string pageRanges) {
        return Split(PdfPageRange.ParseMany(pageRanges));
    }

    /// <summary>
    /// Returns outline/bookmark-derived page ranges in document order.
    /// </summary>
    public IReadOnlyList<PdfBookmarkPageRange> BookmarkPageRanges(params string[] bookmarkTitles) {
        return BookmarkPageRanges(_document.ReadOptions, bookmarkTitles);
    }

    private IReadOnlyList<PdfBookmarkPageRange> BookmarkPageRanges(PdfReadOptions? options, params string[] bookmarkTitles) {
        PdfDocumentInfo info = _document.Inspect(options);
        PdfBookmarkPageRange[] allRanges = BuildBookmarkPageRanges(info);
        if (bookmarkTitles is null || bookmarkTitles.Length == 0) {
            return allRanges;
        }

        var selected = new List<PdfBookmarkPageRange>(bookmarkTitles.Length);
        for (int i = 0; i < bookmarkTitles.Length; i++) {
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

    private PdfDocument[] SplitByBookmarks(PdfReadOptions? options, params string[] bookmarkTitles) {
        IReadOnlyList<PdfBookmarkPageRange> ranges = BookmarkPageRanges(options, bookmarkTitles);
        if (ranges.Count == 0) {
            throw new InvalidOperationException("PDF does not contain any readable bookmarks with page destinations.");
        }

        return Split(ranges.Select(range => range.PageRange), options);
    }

    /// <summary>
    /// Attempts to create one PDF per page, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<PdfDocument>> TrySplit(PdfReadOptions? options = null) {
        return TryPageExtractionOperation<IReadOnlyList<PdfDocument>>("Split pages", effectiveOptions => Split(effectiveOptions), options);
    }

    /// <summary>
    /// Attempts to create PDFs containing consecutive groups of the requested page count.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<PdfDocument>> TrySplit(int pagesPerDocument, PdfReadOptions? options = null) {
        return TryPageExtractionOperation<IReadOnlyList<PdfDocument>>("Split page groups", effectiveOptions => Split(pagesPerDocument, effectiveOptions), options);
    }

    /// <summary>
    /// Attempts to create one PDF for each supplied page selection.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<PdfDocument>> TrySplit(IReadOnlyList<PdfPageSelection> selections, PdfReadOptions? options = null) {
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
    public PdfOperationResult<IReadOnlyList<PdfDocument>> TrySplit(IEnumerable<PdfPageRange> pageRanges, PdfReadOptions? options = null) {
        Guard.NotNull(pageRanges, nameof(pageRanges));
        return TryPageExtractionOperation<IReadOnlyList<PdfDocument>>("Split page ranges", effectiveOptions => Split(pageRanges, effectiveOptions), options);
    }

    /// <summary>
    /// Attempts to create one PDF for each parsed page range, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<PdfDocument>> TrySplit(string pageRanges, PdfReadOptions? options = null) {
        return TryPageExtractionOperation<IReadOnlyList<PdfDocument>>("Split page ranges", effectiveOptions => Split(PdfPageRange.ParseMany(pageRanges), effectiveOptions), options);
    }

    /// <summary>
    /// Attempts to create one PDF for each outline/bookmark-derived page range.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<PdfDocument>> TrySplitByBookmarks(IReadOnlyList<string>? bookmarkTitles = null, PdfReadOptions? options = null) {
        return TryPageExtractionOperation<IReadOnlyList<PdfDocument>>(
            "Split bookmarks",
            effectiveOptions => SplitByBookmarks(effectiveOptions, bookmarkTitles is null ? Array.Empty<string>() : bookmarkTitles.ToArray()),
            options);
    }

    /// <summary>
    /// Creates a new PDF with selected pages deleted.
    /// </summary>
    public PdfDocument Delete(params int[] pageNumbers) {
        return _document.ApplyMutation(input => PdfPageEditor.DeletePages(input, pageNumbers));
    }

    /// <summary>
    /// Creates a new PDF with one inclusive page range deleted.
    /// </summary>
    public PdfDocument Delete(PdfPageRange pageRange) {
        return _document.ApplyMutation(input => PdfPageEditor.DeletePageRange(input, pageRange));
    }

    /// <summary>
    /// Creates a new PDF with selected pages deleted.
    /// </summary>
    public PdfDocument Delete(PdfPageSelection selection) {
        Guard.NotNull(selection, nameof(selection));
        return _document.ApplyMutation(input => PdfPageEditor.DeletePageRanges(input, selection.ToRanges()));
    }

    /// <summary>
    /// Attempts to create a new PDF with selected pages deleted, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryDelete(PdfPageSelection selection, PdfReadOptions? options = null) {
        Guard.NotNull(selection, nameof(selection));
        return _document.TryMutationOperation("Delete pages", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.ModifyPageTree, () => Delete(selection), options);
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
    public PdfOperationResult<PdfDocument> TryDelete(string pageRanges, PdfReadOptions? options = null) {
        return _document.TryMutationOperation("Delete pages", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.ModifyPageTree, () => Delete(PdfPageSelection.Parse(pageRanges)), options);
    }

    /// <summary>
    /// Creates a new PDF with every page copied in the specified one-based order.
    /// </summary>
    public PdfDocument Reorder(params int[] pageNumbers) {
        return _document.ApplyMutation(input => PdfPageEditor.ReorderPages(input, pageNumbers));
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
        return _document.ApplyMutation(input => PdfPageEditor.ReorderPageRanges(input, selection.ToRanges()));
    }

    /// <summary>
    /// Attempts to create a new PDF with every page copied in the selected one-based order, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryReorder(PdfPageSelection selection, PdfReadOptions? options = null) {
        Guard.NotNull(selection, nameof(selection));
        return _document.TryMutationOperation("Reorder pages", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.ModifyPageTree, () => Reorder(selection), options);
    }

    /// <summary>
    /// Attempts to create a new PDF with pages copied in parsed page-range order, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryReorder(string pageRanges, PdfReadOptions? options = null) {
        return _document.TryMutationOperation("Reorder pages", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.ModifyPageTree, () => Reorder(PdfPageSelection.Parse(pageRanges)), options);
    }

    /// <summary>
    /// Creates a new PDF with selected pages duplicated immediately after each source page.
    /// </summary>
    public PdfDocument Duplicate(params int[] pageNumbers) {
        return _document.ApplyMutation(input => PdfPageEditor.DuplicatePages(input, pageNumbers));
    }

    /// <summary>
    /// Creates a new PDF with one inclusive page range duplicated.
    /// </summary>
    public PdfDocument Duplicate(PdfPageRange pageRange) {
        return _document.ApplyMutation(input => PdfPageEditor.DuplicatePageRange(input, pageRange));
    }

    /// <summary>
    /// Creates a new PDF with selected pages duplicated immediately after each source page.
    /// </summary>
    public PdfDocument Duplicate(PdfPageSelection selection) {
        Guard.NotNull(selection, nameof(selection));
        return _document.ApplyMutation(input => PdfPageEditor.DuplicatePageRanges(input, selection.ToRanges()));
    }

    /// <summary>
    /// Attempts to create a new PDF with selected pages duplicated, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryDuplicate(PdfPageSelection selection, PdfReadOptions? options = null) {
        Guard.NotNull(selection, nameof(selection));
        return _document.TryMutationOperation("Duplicate pages", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.ModifyPageTree, () => Duplicate(selection), options);
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
    public PdfOperationResult<PdfDocument> TryDuplicate(string pageRanges, PdfReadOptions? options = null) {
        return _document.TryMutationOperation("Duplicate pages", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.ModifyPageTree, () => Duplicate(PdfPageSelection.Parse(pageRanges)), options);
    }

    /// <summary>
    /// Creates a new PDF with selected pages moved before the supplied one-based page number.
    /// Use page count + 1 to move pages to the end.
    /// </summary>
    public PdfDocument Move(int insertBeforePageNumber, params int[] pageNumbers) {
        return _document.ApplyMutation(input => PdfPageEditor.MovePages(input, insertBeforePageNumber, pageNumbers));
    }

    /// <summary>
    /// Creates a new PDF with one inclusive page range moved before the supplied one-based page number.
    /// Use page count + 1 to move pages to the end.
    /// </summary>
    public PdfDocument Move(int insertBeforePageNumber, PdfPageRange pageRange) {
        return _document.ApplyMutation(input => PdfPageEditor.MovePageRange(input, insertBeforePageNumber, pageRange));
    }

    /// <summary>
    /// Creates a new PDF with selected pages moved before the supplied one-based page number.
    /// Use page count + 1 to move pages to the end.
    /// </summary>
    public PdfDocument Move(int insertBeforePageNumber, PdfPageSelection selection) {
        Guard.NotNull(selection, nameof(selection));
        return _document.ApplyMutation(input => PdfPageEditor.MovePageRanges(input, insertBeforePageNumber, selection.ToRanges()));
    }

    /// <summary>
    /// Attempts to create a new PDF with selected pages moved before the supplied one-based page number, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryMove(int insertBeforePageNumber, PdfPageSelection selection, PdfReadOptions? options = null) {
        Guard.NotNull(selection, nameof(selection));
        return _document.TryMutationOperation("Move pages", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.ModifyPageTree, () => Move(insertBeforePageNumber, selection), options);
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
    public PdfOperationResult<PdfDocument> TryMove(int insertBeforePageNumber, string pageRanges, PdfReadOptions? options = null) {
        return _document.TryMutationOperation("Move pages", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.ModifyPageTree, () => Move(insertBeforePageNumber, PdfPageSelection.Parse(pageRanges)), options);
    }

    /// <summary>
    /// Creates a new PDF with selected pages rotated. Supplying no page numbers rotates every page.
    /// </summary>
    public PdfDocument Rotate(int rotationDegrees, params int[] pageNumbers) {
        return _document.ApplyMutation(input => PdfPageEditor.RotatePages(input, rotationDegrees, pageNumbers));
    }

    /// <summary>
    /// Creates a new PDF with one inclusive page range rotated.
    /// </summary>
    public PdfDocument Rotate(int rotationDegrees, PdfPageRange pageRange) {
        return _document.ApplyMutation(input => PdfPageEditor.RotatePageRange(input, rotationDegrees, pageRange));
    }

    /// <summary>
    /// Creates a new PDF with selected pages rotated.
    /// </summary>
    public PdfDocument Rotate(int rotationDegrees, PdfPageSelection selection) {
        Guard.NotNull(selection, nameof(selection));
        return _document.ApplyMutation(input => PdfPageEditor.RotatePageRanges(input, rotationDegrees, selection.ToRanges()));
    }

    /// <summary>
    /// Attempts to create a new PDF with selected pages rotated, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<PdfDocument> TryRotate(int rotationDegrees, PdfPageSelection selection, PdfReadOptions? options = null) {
        Guard.NotNull(selection, nameof(selection));
        return _document.TryMutationOperation("Rotate pages", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.ModifyPageTree, () => Rotate(rotationDegrees, selection), options);
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
    public PdfOperationResult<PdfDocument> TryRotate(int rotationDegrees, string pageRanges, PdfReadOptions? options = null) {
        return _document.TryMutationOperation("Rotate pages", PdfPreflightCapability.ManipulatePages, PdfMutationOperation.ModifyPageTree, () => Rotate(rotationDegrees, PdfPageSelection.Parse(pageRanges)), options);
    }

    /// <summary>Sets a typed boundary box for selected pages, or every page when no page numbers are supplied.</summary>
    public PdfDocument SetPageBox(PdfPageBoundaryBox box, double left, double bottom, double right, double top, params int[] pageNumbers) {
        return _document.ApplyMutation(input => PdfPageEditor.SetPageBox(input, box, left, bottom, right, top, pageNumbers));
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
            input, left, bottom, right, top, pageNumbers));

    /// <summary>Destructively crops selected pages by replacing the retained visual rectangle with a validated opaque raster page.</summary>
    public PdfDestructiveCropResult DestructiveCrop(double left, double bottom, double right, double top, PdfDestructiveCropOptions? options = null, params int[] pageNumbers) =>
        PdfPageEditor.DestructiveCropPages(_document.GetBytesForOperation(), left, bottom, right, top, options, pageNumbers);

    /// <summary>Sets the bleed box for selected pages.</summary>
    public PdfDocument SetBleedBox(double left, double bottom, double right, double top, params int[] pageNumbers) =>
        SetPageBox(PdfPageBoundaryBox.BleedBox, left, bottom, right, top, pageNumbers);

    /// <summary>Sets the trim box for selected pages.</summary>
    public PdfDocument SetTrimBox(double left, double bottom, double right, double top, params int[] pageNumbers) =>
        SetPageBox(PdfPageBoundaryBox.TrimBox, left, bottom, right, top, pageNumbers);

    /// <summary>Sets the art box for selected pages.</summary>
    public PdfDocument SetArtBox(double left, double bottom, double right, double top, params int[] pageNumbers) =>
        SetPageBox(PdfPageBoundaryBox.ArtBox, left, bottom, right, top, pageNumbers);

    private static PdfBookmarkPageRange[] BuildBookmarkPageRanges(PdfDocumentInfo info) {
        var outlines = new List<PdfOutlineItem>();
        FlattenOutlines(info.Outlines, outlines);

        var anchors = new List<PdfOutlineItem>();
        for (int i = 0; i < outlines.Count; i++) {
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

    private static void FlattenOutlines(IReadOnlyList<PdfOutlineItem> source, List<PdfOutlineItem> destination) {
        for (int i = 0; i < source.Count; i++) {
            PdfOutlineItem outline = source[i];
            destination.Add(outline);
            if (outline.Children.Count > 0) {
                FlattenOutlines(outline.Children, destination);
            }
        }
    }

    private PdfOperationResult<T> TryPageExtractionOperation<T>(
        string operationName,
        Func<PdfReadOptions?, T> operation,
        PdfReadOptions? options) where T : class {
        PdfReadOptions? effectiveOptions = options ?? _document.ReadOptions;
        PdfMutationPlan plan = _document.PlanMutation(PdfMutationOperation.ExtractPages, options: effectiveOptions);
        if (!plan.CanExecute) {
            return PdfOperationResult<T>.MutationBlocked(operationName, PdfPreflightCapability.ManipulatePages, plan);
        }

        try {
            return PdfOperationResult<T>.MutationSuccess(
                operationName,
                PdfPreflightCapability.ManipulatePages,
                plan,
                operation(effectiveOptions));
        } catch (Exception ex) {
            return PdfOperationResult<T>.MutationFailed(
                operationName,
                PdfPreflightCapability.ManipulatePages,
                plan,
                ex);
        }
    }
}
