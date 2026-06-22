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
    /// Creates PDFs containing consecutive groups of the requested page count.
    /// </summary>
    public IReadOnlyList<PdfDocument> Split(int pagesPerDocument) {
        if (pagesPerDocument <= 0) {
            throw new ArgumentOutOfRangeException(nameof(pagesPerDocument), pagesPerDocument, "Pages per document must be greater than zero.");
        }

        int pageCount = _document.Inspect().PageCount;
        if (pageCount == 0) {
            throw new InvalidOperationException("PDF does not contain any readable pages.");
        }

        var ranges = new List<PdfPageRange>();
        for (int firstPage = 1; firstPage <= pageCount; firstPage += pagesPerDocument) {
            int lastPage = Math.Min(firstPage + pagesPerDocument - 1, pageCount);
            ranges.Add(PdfPageRange.From(firstPage, lastPage));
        }

        return Split(ranges);
    }

    /// <summary>
    /// Creates one PDF for each supplied page selection.
    /// </summary>
    public IReadOnlyList<PdfDocument> Split(params PdfPageSelection[] selections) {
        Guard.NotNull(selections, nameof(selections));
        if (selections.Length == 0) {
            throw new ArgumentException("At least one page selection must be specified.", nameof(selections));
        }

        var documents = new PdfDocument[selections.Length];
        for (int i = 0; i < selections.Length; i++) {
            Guard.NotNull(selections[i], nameof(selections));
            documents[i] = Extract(selections[i]);
        }

        return documents;
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

        return PdfPageExtractor.SplitPageRanges(_document.Snapshot(), ranges)
            .Select(PdfDocument.FromBytes)
            .ToArray();
    }

    /// <summary>
    /// Returns outline/bookmark-derived page ranges in document order.
    /// </summary>
    public IReadOnlyList<PdfBookmarkPageRange> BookmarkPageRanges(params string[] bookmarkTitles) {
        PdfDocumentInfo info = _document.Inspect();
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

    /// <summary>
    /// Attempts to create one PDF per page, returning diagnostics when blocked or failed.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<PdfDocument>> TrySplit(PdfReadOptions? options = null) {
        return _document.TryOperation("Split pages", PdfPreflightCapability.ManipulatePages, Split, options);
    }

    /// <summary>
    /// Attempts to create PDFs containing consecutive groups of the requested page count.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<PdfDocument>> TrySplit(int pagesPerDocument, PdfReadOptions? options = null) {
        return _document.TryOperation("Split page groups", PdfPreflightCapability.ManipulatePages, () => Split(pagesPerDocument), options);
    }

    /// <summary>
    /// Attempts to create one PDF for each supplied page selection.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<PdfDocument>> TrySplit(IReadOnlyList<PdfPageSelection> selections, PdfReadOptions? options = null) {
        Guard.NotNull(selections, nameof(selections));
        return _document.TryOperation("Split page selections", PdfPreflightCapability.ManipulatePages, () => Split(selections.ToArray()), options);
    }

    /// <summary>
    /// Attempts to create one PDF for each outline/bookmark-derived page range.
    /// </summary>
    public PdfOperationResult<IReadOnlyList<PdfDocument>> TrySplitByBookmarks(IReadOnlyList<string>? bookmarkTitles = null, PdfReadOptions? options = null) {
        return _document.TryOperation(
            "Split bookmarks",
            PdfPreflightCapability.ManipulatePages,
            () => SplitByBookmarks(bookmarkTitles is null ? Array.Empty<string>() : bookmarkTitles.ToArray()),
            options);
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
}
