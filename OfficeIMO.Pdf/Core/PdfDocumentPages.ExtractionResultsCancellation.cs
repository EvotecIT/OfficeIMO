using System.Threading;

namespace OfficeIMO.Pdf;

public sealed partial class PdfDocumentPages {
    /// <summary>Extracts pages with diagnostics while preserving cancellation as an exception.</summary>
    public PdfOperationResult<PdfDocument> ExtractResult(PdfPageSelection selection, CancellationToken cancellationToken, PdfLoadOptions? options = null) {
        Guard.NotNull(selection, nameof(selection));
        return TryPageExtractionOperation("Extract pages", effective => Extract(selection, effective, cancellationToken), options, cancellationToken);
    }

    /// <summary>Extracts pages with diagnostics while preserving cancellation as an exception.</summary>
    public PdfOperationResult<PdfDocument> ExtractResult(PdfPageSelector selector, CancellationToken cancellationToken, PdfLoadOptions? options = null) {
        Guard.NotNull(selector, nameof(selector));
        return TryPageExtractionOperation("Extract pages", effective => Extract(ResolveSelector(selector, effective, cancellationToken), effective, cancellationToken), options, cancellationToken);
    }

    /// <summary>Extracts pages with diagnostics while preserving cancellation as an exception.</summary>
    public PdfOperationResult<PdfDocument> ExtractResult(string pageRanges, CancellationToken cancellationToken, PdfLoadOptions? options = null) =>
        TryPageExtractionOperation("Extract pages", effective => Extract(PdfPageSelection.Parse(pageRanges), effective, cancellationToken), options, cancellationToken);

    /// <summary>Creates split results while propagating cancellation through every generated part.</summary>
    public PdfOperationResult<IReadOnlyList<PdfDocument>> SplitResult(CancellationToken cancellationToken, PdfLoadOptions? options = null) =>
        TryPageExtractionOperation<IReadOnlyList<PdfDocument>>("Split pages", effective => Split(effective, cancellationToken), options, cancellationToken);

    /// <summary>Creates split results while propagating cancellation through every generated part.</summary>
    public PdfOperationResult<IReadOnlyList<PdfDocument>> SplitResult(int pagesPerDocument, CancellationToken cancellationToken, PdfLoadOptions? options = null) =>
        TryPageExtractionOperation<IReadOnlyList<PdfDocument>>("Split page groups", effective => Split(pagesPerDocument, effective, cancellationToken), options, cancellationToken);

    /// <summary>Creates split results while propagating cancellation through every generated part.</summary>
    public PdfOperationResult<IReadOnlyList<PdfDocument>> SplitResult(IReadOnlyList<PdfPageSelection> selections, CancellationToken cancellationToken, PdfLoadOptions? options = null) {
        Guard.NotNull(selections, nameof(selections));
        if (selections.Count == 0) {
            return TryPageExtractionOperation<IReadOnlyList<PdfDocument>>("Split page selections",
                _ => throw new ArgumentException("At least one page selection must be specified.", nameof(selections)), options, cancellationToken);
        }
        return TryPageExtractionOperation<IReadOnlyList<PdfDocument>>("Split page selections", effective => Split(selections.ToArray(), effective, cancellationToken), options, cancellationToken);
    }

    /// <summary>Creates split results while propagating cancellation through every generated part.</summary>
    public PdfOperationResult<IReadOnlyList<PdfDocument>> SplitResult(IEnumerable<PdfPageRange> pageRanges, CancellationToken cancellationToken, PdfLoadOptions? options = null) {
        Guard.NotNull(pageRanges, nameof(pageRanges));
        return TryPageExtractionOperation<IReadOnlyList<PdfDocument>>("Split page ranges", effective => Split(pageRanges, effective, cancellationToken), options, cancellationToken);
    }

    /// <summary>Creates split results while propagating cancellation through every generated part.</summary>
    public PdfOperationResult<IReadOnlyList<PdfDocument>> SplitResult(string pageRanges, CancellationToken cancellationToken, PdfLoadOptions? options = null) =>
        TryPageExtractionOperation<IReadOnlyList<PdfDocument>>("Split page ranges", effective => Split(PdfPageRange.ParseMany(pageRanges), effective, cancellationToken), options, cancellationToken);

    /// <summary>Resolves bookmark ranges while observing cancellation and retaining captured read limits.</summary>
    public IReadOnlyList<PdfBookmarkPageRange> BookmarkPageRanges(CancellationToken cancellationToken, params string[] bookmarkTitles) =>
        BookmarkPageRanges(_document.ReadOptions, cancellationToken, bookmarkTitles);

    /// <summary>Splits bookmarked sections with cooperative cancellation through every output part.</summary>
    public IReadOnlyList<PdfDocument> SplitByBookmarks(CancellationToken cancellationToken, params string[] bookmarkTitles) =>
        SplitByBookmarks(_document.ReadOptions, cancellationToken, bookmarkTitles);

    /// <summary>Reports bookmark splitting while preserving cancellation as an exception.</summary>
    public PdfOperationResult<IReadOnlyList<PdfDocument>> SplitByBookmarksResult(CancellationToken cancellationToken, IReadOnlyList<string>? bookmarkTitles = null, PdfLoadOptions? options = null) =>
        TryPageExtractionOperation<IReadOnlyList<PdfDocument>>("Split bookmarks", effective => SplitByBookmarks(effective, cancellationToken, bookmarkTitles?.ToArray() ?? Array.Empty<string>()), options, cancellationToken);
}
