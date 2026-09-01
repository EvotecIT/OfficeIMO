namespace OfficeIMO.Pdf;

/// <summary>Fluent cross-page paragraph recovery operations for a <see cref="PdfDocument"/>.</summary>
internal sealed partial class PdfDocumentReader {
    /// <summary>Recovers conservative cross-page paragraph continuation groups.</summary>
    public IReadOnlyList<PdfLogicalParagraphContinuationGroup> ParagraphContinuations(
        PdfLogicalParagraphContinuationOptions? continuationOptions = null,
        PdfTextLayoutOptions? layoutOptions = null,
        PdfLoadOptions? readOptions = null) {
        return Logical(layoutOptions, readOptions).GetParagraphContinuationGroups(continuationOptions);
    }

    /// <summary>Attempts to recover cross-page paragraph continuation groups.</summary>
    public PdfOperationResult<IReadOnlyList<PdfLogicalParagraphContinuationGroup>> TryParagraphContinuations(
        PdfLogicalParagraphContinuationOptions? continuationOptions = null,
        PdfTextLayoutOptions? layoutOptions = null,
        PdfLoadOptions? readOptions = null) {
        return _document.TryOperation(
            "Recover paragraph continuations",
            PdfPreflightCapability.ReadLogicalObjects,
            () => ParagraphContinuations(continuationOptions, layoutOptions, readOptions),
            ResolveReadOptions(readOptions));
    }

    /// <summary>Recovers conservative cross-page paragraph continuation groups for selected pages.</summary>
    public IReadOnlyList<PdfLogicalParagraphContinuationGroup> ParagraphContinuations(
        PdfPageSelection selection,
        PdfLogicalParagraphContinuationOptions? continuationOptions = null,
        PdfTextLayoutOptions? layoutOptions = null,
        PdfLoadOptions? readOptions = null) {
        Guard.NotNull(selection, nameof(selection));
        return Logical(selection, layoutOptions, readOptions).GetParagraphContinuationGroups(continuationOptions);
    }

    /// <summary>Attempts to recover cross-page paragraph continuation groups for selected pages.</summary>
    public PdfOperationResult<IReadOnlyList<PdfLogicalParagraphContinuationGroup>> TryParagraphContinuations(
        PdfPageSelection selection,
        PdfLogicalParagraphContinuationOptions? continuationOptions = null,
        PdfTextLayoutOptions? layoutOptions = null,
        PdfLoadOptions? readOptions = null) {
        Guard.NotNull(selection, nameof(selection));
        return _document.TryOperation(
            "Recover paragraph continuations",
            PdfPreflightCapability.ReadLogicalObjects,
            () => ParagraphContinuations(selection, continuationOptions, layoutOptions, readOptions),
            ResolveReadOptions(readOptions));
    }
}
