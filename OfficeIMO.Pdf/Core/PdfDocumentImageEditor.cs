namespace OfficeIMO.Pdf;

/// <summary>Finds and edits image placements on existing PDF pages.</summary>
public sealed class PdfDocumentImageEditor {
    private readonly PdfDocument _document;

    internal PdfDocumentImageEditor(PdfDocument document) => _document = document;

    /// <summary>Extracts every unique embedded image XObject in page order.</summary>
    public IReadOnlyList<PdfExtractedImage> Extract(PdfLoadOptions? readOptions = null) =>
        _document.Reader.Images(readOptions);

    /// <summary>Extracts embedded images referenced by a caller-ordered page selection.</summary>
    public IReadOnlyList<PdfExtractedImage> Extract(PdfPageSelection selection, PdfLoadOptions? readOptions = null) =>
        _document.Reader.Images(selection, readOptions);

    /// <summary>Extracts embedded images referenced by parsed one-based page ranges.</summary>
    public IReadOnlyList<PdfExtractedImage> Extract(string pageRanges, PdfLoadOptions? readOptions = null) =>
        _document.Reader.Images(pageRanges, readOptions);

    /// <summary>Extracts embedded images referenced by a document-relative page selector.</summary>
    public IReadOnlyList<PdfExtractedImage> Extract(PdfPageSelector selector, PdfLoadOptions? readOptions = null) =>
        _document.Reader.Images(selector, readOptions);

    /// <summary>Attempts to extract images, returning preflight diagnostics when blocked or failed.</summary>
    public PdfOperationResult<IReadOnlyList<PdfExtractedImage>> ExtractResult(PdfLoadOptions? readOptions = null) =>
        _document.Reader.ImagesResult(readOptions);

    /// <summary>Returns every image placement invocation in page paint order.</summary>
    public IReadOnlyList<PdfImagePlacement> Placements(PdfLoadOptions? readOptions = null) =>
        PdfImageEditor.Placements(_document.GetBytesForOperation(), readOptions ?? _document.ReadOptions);

    /// <summary>Returns image placements referenced by a caller-ordered page selection.</summary>
    public IReadOnlyList<PdfImagePlacement> Placements(PdfPageSelection selection, PdfLoadOptions? readOptions = null) =>
        SelectPlacements(selection, readOptions);

    /// <summary>Returns image placements referenced by parsed one-based page ranges.</summary>
    public IReadOnlyList<PdfImagePlacement> Placements(string pageRanges, PdfLoadOptions? readOptions = null) =>
        SelectPlacements(PdfPageSelection.Parse(pageRanges), readOptions);

    /// <summary>Returns image placements referenced by a document-relative page selector.</summary>
    public IReadOnlyList<PdfImagePlacement> Placements(PdfPageSelector selector, PdfLoadOptions? readOptions = null) {
        Guard.NotNull(selector, nameof(selector));
        PdfLoadOptions effective = readOptions ?? _document.ReadOptions;
        int pageCount = _document.GetReadDocument(effective).Pages.Count;
        if (pageCount == 0) throw new InvalidOperationException("PDF does not contain any readable pages.");
        return SelectPlacements(selector.ResolveSelection(pageCount), effective);
    }

    /// <summary>Attempts to inspect image placements, returning preflight diagnostics when blocked or failed.</summary>
    public PdfOperationResult<IReadOnlyList<PdfImagePlacement>> PlacementsResult(PdfLoadOptions? readOptions = null) =>
        _document.TryOperation(
            "Extract image placements",
            PdfPreflightCapability.ExtractImages,
            () => Placements(readOptions),
            readOptions ?? _document.ReadOptions);

    /// <summary>Returns image placements whose page-space bounds intersect a region.</summary>
    public IReadOnlyList<PdfImagePlacement> Find(PdfPageRegion region, PdfLoadOptions? readOptions = null) =>
        PdfImageEditor.Find(_document.GetBytesForOperation(), region, readOptions ?? _document.ReadOptions);

    /// <summary>Adds an image fitted to a page region.</summary>
    public PdfImageEditResult Add(PdfPageRegion target, byte[] imageBytes, PdfImageEditOptions? options = null, PdfLoadOptions? readOptions = null) =>
        Apply(input => PdfImageEditor.Add(input, target, imageBytes, options, readOptions ?? _document.ReadOptions), readOptions);

    /// <summary>Removes one exact image placement without removing text, paths, annotations, or other overlapping images.</summary>
    public PdfImageEditResult Remove(PdfImagePlacement placement, PdfLoadOptions? readOptions = null) =>
        Apply(input => PdfImageEditor.Remove(input, placement, readOptions ?? _document.ReadOptions), readOptions);

    /// <summary>
    /// Replaces one exact image placement while preserving portable position, size, and rotation. The selected
    /// <see cref="PdfImageEditOptions.Layer"/> controls its new relation to existing page content.
    /// </summary>
    public PdfImageEditResult Replace(PdfImagePlacement placement, byte[] imageBytes, PdfImageEditOptions? options = null, PdfLoadOptions? readOptions = null) =>
        Apply(input => PdfImageEditor.Replace(input, placement, imageBytes, options, readOptions ?? _document.ReadOptions), readOptions);

    /// <summary>
    /// Moves one exact image placement by a PDF user-space offset. The source payload, portable size, and rotation
    /// are preserved; the selected <see cref="PdfImageEditOptions.Layer"/> controls its new page-content layer.
    /// </summary>
    public PdfImageEditResult Move(PdfImagePlacement placement, double deltaX, double deltaY, PdfImageEditOptions? options = null, PdfLoadOptions? readOptions = null) =>
        Apply(input => PdfImageEditor.Move(input, placement, deltaX, deltaY, options, readOptions ?? _document.ReadOptions), readOptions);

    private PdfImageEditResult Apply(Func<byte[], PdfImageEditor.ImageMutationResult> operation, PdfLoadOptions? readOptions) {
        PdfImageEditor.ImageMutationResult? mutation = null;
        PdfDocument document = _document.ApplyMutation(input => {
            mutation = operation(input);
            return mutation.Bytes;
        }, readOptions, operationName: "Image");
        if (mutation is null) throw new InvalidOperationException("PDF image edit did not produce a mutation result.");
        return new PdfImageEditResult(document, mutation.AffectedCount);
    }

    private System.Collections.ObjectModel.ReadOnlyCollection<PdfImagePlacement> SelectPlacements(
        PdfPageSelection selection,
        PdfLoadOptions? readOptions) {
        Guard.NotNull(selection, nameof(selection));
        PdfLoadOptions effective = readOptions ?? _document.ReadOptions;
        int pageCount = _document.GetReadDocument(effective).Pages.Count;
        int[] selectedPages = selection.ToPageNumbers(pageCount, nameof(selection));
        IReadOnlyList<PdfImagePlacement> placements = Placements(effective);
        var selected = new List<PdfImagePlacement>();
        for (int pageIndex = 0; pageIndex < selectedPages.Length; pageIndex++) {
            int pageNumber = selectedPages[pageIndex];
            for (int placementIndex = 0; placementIndex < placements.Count; placementIndex++) {
                PdfImagePlacement placement = placements[placementIndex];
                if (placement.PageNumber == pageNumber) selected.Add(placement);
            }
        }
        return selected.AsReadOnly();
    }
}
