namespace OfficeIMO.Pdf;

/// <summary>Finds and edits image placements on existing PDF pages.</summary>
public sealed class PdfDocumentImageEditor {
    private readonly PdfDocument _document;

    internal PdfDocumentImageEditor(PdfDocument document) => _document = document;

    /// <summary>Returns every image placement invocation in page paint order.</summary>
    public IReadOnlyList<PdfImagePlacement> Placements(PdfReadOptions? readOptions = null) =>
        PdfImageEditor.Placements(_document.GetBytesForOperation(), readOptions ?? _document.ReadOptions);

    /// <summary>Returns image placements whose page-space bounds intersect a region.</summary>
    public IReadOnlyList<PdfImagePlacement> Find(PdfPageRegion region, PdfReadOptions? readOptions = null) =>
        PdfImageEditor.Find(_document.GetBytesForOperation(), region, readOptions ?? _document.ReadOptions);

    /// <summary>Adds an image fitted to a page region.</summary>
    public PdfImageEditResult Add(PdfPageRegion target, byte[] imageBytes, PdfImageEditOptions? options = null, PdfReadOptions? readOptions = null) =>
        Apply(input => PdfImageEditor.Add(input, target, imageBytes, options, readOptions ?? _document.ReadOptions), readOptions);

    /// <summary>Removes one exact image placement without removing text, paths, annotations, or other overlapping images.</summary>
    public PdfImageEditResult Remove(PdfImagePlacement placement, PdfReadOptions? readOptions = null) =>
        Apply(input => PdfImageEditor.Remove(input, placement, readOptions ?? _document.ReadOptions), readOptions);

    /// <summary>
    /// Replaces one exact image placement while preserving portable position, size, and rotation. The selected
    /// <see cref="PdfImageEditOptions.Layer"/> controls its new relation to existing page content.
    /// </summary>
    public PdfImageEditResult Replace(PdfImagePlacement placement, byte[] imageBytes, PdfImageEditOptions? options = null, PdfReadOptions? readOptions = null) =>
        Apply(input => PdfImageEditor.Replace(input, placement, imageBytes, options, readOptions ?? _document.ReadOptions), readOptions);

    /// <summary>
    /// Moves one exact image placement by a PDF user-space offset. The source payload, portable size, and rotation
    /// are preserved; the selected <see cref="PdfImageEditOptions.Layer"/> controls its new page-content layer.
    /// </summary>
    public PdfImageEditResult Move(PdfImagePlacement placement, double deltaX, double deltaY, PdfImageEditOptions? options = null, PdfReadOptions? readOptions = null) =>
        Apply(input => PdfImageEditor.Move(input, placement, deltaX, deltaY, options, readOptions ?? _document.ReadOptions), readOptions);

    private PdfImageEditResult Apply(Func<byte[], PdfImageEditor.ImageMutationResult> operation, PdfReadOptions? readOptions) {
        PdfImageEditor.ImageMutationResult? mutation = null;
        PdfDocument document = _document.ApplyMutation(input => {
            mutation = operation(input);
            return mutation.Bytes;
        }, readOptions, operationName: "Image");
        if (mutation is null) throw new InvalidOperationException("PDF image edit did not produce a mutation result.");
        return new PdfImageEditResult(document, mutation.AffectedCount);
    }
}
