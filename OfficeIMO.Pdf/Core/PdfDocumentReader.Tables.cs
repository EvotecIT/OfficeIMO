namespace OfficeIMO.Pdf;

/// <summary>Fluent cross-page table recovery operations for a <see cref="PdfDocument"/>.</summary>
public sealed partial class PdfDocumentReader {
    /// <summary>Recovers bounded table continuation groups for the full document.</summary>
    public IReadOnlyList<PdfLogicalTableContinuationGroup> TableContinuations(
        PdfLogicalTableContinuationOptions? continuationOptions = null,
        PdfTextLayoutOptions? layoutOptions = null,
        PdfReadOptions? readOptions = null) =>
        Logical(layoutOptions, readOptions).GetTableContinuationGroups(continuationOptions);

    /// <summary>Attempts to recover table continuation groups for the full document.</summary>
    public PdfOperationResult<IReadOnlyList<PdfLogicalTableContinuationGroup>> TryTableContinuations(
        PdfLogicalTableContinuationOptions? continuationOptions = null,
        PdfTextLayoutOptions? layoutOptions = null,
        PdfReadOptions? readOptions = null) =>
        _document.TryOperation(
            "Recover table continuations",
            PdfPreflightCapability.ReadLogicalObjects,
            () => TableContinuations(continuationOptions, layoutOptions, readOptions),
            ResolveReadOptions(readOptions));

    /// <summary>Recovers bounded table continuation groups for selected pages.</summary>
    public IReadOnlyList<PdfLogicalTableContinuationGroup> TableContinuations(
        PdfPageSelection selection,
        PdfLogicalTableContinuationOptions? continuationOptions = null,
        PdfTextLayoutOptions? layoutOptions = null,
        PdfReadOptions? readOptions = null) {
        Guard.NotNull(selection, nameof(selection));
        return Logical(selection, layoutOptions, readOptions).GetTableContinuationGroups(continuationOptions);
    }

    /// <summary>Attempts to recover table continuation groups for selected pages.</summary>
    public PdfOperationResult<IReadOnlyList<PdfLogicalTableContinuationGroup>> TryTableContinuations(
        PdfPageSelection selection,
        PdfLogicalTableContinuationOptions? continuationOptions = null,
        PdfTextLayoutOptions? layoutOptions = null,
        PdfReadOptions? readOptions = null) {
        Guard.NotNull(selection, nameof(selection));
        return _document.TryOperation(
            "Recover table continuations",
            PdfPreflightCapability.ReadLogicalObjects,
            () => TableContinuations(selection, continuationOptions, layoutOptions, readOptions),
            ResolveReadOptions(readOptions));
    }

    /// <summary>Recovers bounded table continuation groups for one-based page ranges.</summary>
    public IReadOnlyList<PdfLogicalTableContinuationGroup> TableContinuations(
        string pageRanges,
        PdfLogicalTableContinuationOptions? continuationOptions = null,
        PdfTextLayoutOptions? layoutOptions = null,
        PdfReadOptions? readOptions = null) =>
        TableContinuations(PdfPageSelection.Parse(pageRanges), continuationOptions, layoutOptions, readOptions);
}
