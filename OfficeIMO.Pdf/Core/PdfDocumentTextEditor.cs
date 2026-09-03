namespace OfficeIMO.Pdf;

/// <summary>Searches and edits text on existing PDF pages.</summary>
public sealed class PdfDocumentTextEditor {
    private readonly PdfDocument _document;

    internal PdfDocumentTextEditor(PdfDocument document) => _document = document;

    /// <summary>Returns text and dominant style inside a page region.</summary>
    public PdfRegionText Inspect(PdfPageRegion region, PdfLoadOptions? readOptions = null) =>
        PdfTextEditor.Inspect(_document.GetBytesForOperation(), region, readOptions ?? _document.ReadOptions);

    /// <summary>Finds visible, unclipped text occurrences and optionally OCR-style rendering-mode-3 text.</summary>
    public IReadOnlyList<PdfTextMatch> Find(string text, PdfTextSearchOptions? options = null, PdfLoadOptions? readOptions = null) =>
        PdfTextEditor.Find(_document.GetBytesForOperation(), text, options, readOptions ?? _document.ReadOptions);

    /// <summary>Adds text at the top-left of a page region without removing existing content.</summary>
    public PdfTextEditResult Add(PdfPageRegion region, string text, PdfTextEditOptions? options = null, PdfLoadOptions? readOptions = null) {
        PdfLoadOptions? effectiveReadOptions = readOptions ?? _document.ReadOptions;
        return Apply(input => PdfTextEditor.Add(input, region, text, options, effectiveReadOptions), effectiveReadOptions);
    }

    /// <summary>Removes text objects intersecting a region and adds replacement text in the detected style.</summary>
    public PdfTextEditResult Replace(PdfPageRegion region, string text, PdfTextEditOptions? options = null, PdfLoadOptions? readOptions = null) {
        PdfLoadOptions? effectiveReadOptions = readOptions ?? _document.ReadOptions;
        return Apply(input => PdfTextEditor.Replace(input, region, text, options, effectiveReadOptions), effectiveReadOptions);
    }

    /// <summary>Replaces one located text occurrence while preserving unmatched source-span text.</summary>
    public PdfTextEditResult Replace(PdfTextMatch match, string text, PdfTextEditOptions? options = null, PdfLoadOptions? readOptions = null) {
        PdfLoadOptions? effectiveReadOptions = readOptions ?? _document.ReadOptions;
        return Apply(input => PdfTextEditor.Replace(input, match, text, options, effectiveReadOptions), effectiveReadOptions);
    }

    /// <summary>Moves text objects intersecting a region by a PDF user-space offset.</summary>
    public PdfTextEditResult Move(PdfPageRegion region, double deltaX, double deltaY, PdfTextEditOptions? options = null, PdfLoadOptions? readOptions = null) {
        PdfLoadOptions? effectiveReadOptions = readOptions ?? _document.ReadOptions;
        return Apply(input => PdfTextEditor.Move(input, region, deltaX, deltaY, options, effectiveReadOptions), effectiveReadOptions);
    }

    /// <summary>Moves one located text occurrence while leaving unmatched source-span text at its original position.</summary>
    public PdfTextEditResult Move(PdfTextMatch match, double deltaX, double deltaY, PdfTextEditOptions? options = null, PdfLoadOptions? readOptions = null) {
        PdfLoadOptions? effectiveReadOptions = readOptions ?? _document.ReadOptions;
        return Apply(input => PdfTextEditor.Move(input, match, deltaX, deltaY, options, effectiveReadOptions), effectiveReadOptions);
    }

    /// <summary>Replaces every located occurrence while preserving exact unmatched source-span text and independent visual runs.</summary>
    public PdfTextEditResult ReplaceAll(string text, string replacement, PdfTextSearchOptions? searchOptions = null, PdfTextEditOptions? editOptions = null, PdfLoadOptions? readOptions = null) {
        PdfLoadOptions? effectiveReadOptions = readOptions ?? _document.ReadOptions;
        return Apply(input => PdfTextEditor.ReplaceAll(input, text, replacement, searchOptions, editOptions, effectiveReadOptions), effectiveReadOptions);
    }

    private PdfTextEditResult Apply(Func<byte[], PdfTextEditor.TextMutationResult> operation, PdfLoadOptions? readOptions) {
        PdfTextEditor.TextMutationResult? mutation = null;
        PdfDocument document = _document.ApplyMutation(input => {
            mutation = operation(input);
            return mutation.Bytes;
        }, readOptions, operationName: "Text");
        if (mutation is null) throw new InvalidOperationException("PDF text edit did not produce a mutation result.");
        return new PdfTextEditResult(document, mutation.AffectedCount, mutation.Warnings);
    }
}
