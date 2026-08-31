namespace OfficeIMO.Pdf;

/// <summary>Bounded inspection of low-level resources owned by an existing PDF.</summary>
public sealed class PdfDocumentResources {
    private readonly PdfDocument _document;

    internal PdfDocumentResources(PdfDocument document) => _document = document;

    /// <summary>Inspects unique fonts declared by all pages and nested Form XObjects.</summary>
    public PdfFontInventory Fonts(
        PdfFontInspectionOptions? inspectionOptions = null,
        PdfLoadOptions? loadOptions = null) =>
        PdfFontInspector.Inspect(_document.GetReadDocument(loadOptions ?? _document.ReadOptions), inspectionOptions);

    /// <summary>Inspects unique fonts declared by selected pages and nested Form XObjects.</summary>
    public PdfFontInventory Fonts(
        PdfPageSelection selection,
        PdfFontInspectionOptions? inspectionOptions = null,
        PdfLoadOptions? loadOptions = null) {
        Guard.NotNull(selection, nameof(selection));
        return PdfFontInspector.Inspect(
            _document.GetReadDocument(loadOptions ?? _document.ReadOptions),
            inspectionOptions,
            selection);
    }

    /// <summary>Inspects unique fonts declared by pages described by one-based page ranges.</summary>
    public PdfFontInventory Fonts(
        string pageRanges,
        PdfFontInspectionOptions? inspectionOptions = null,
        PdfLoadOptions? loadOptions = null) =>
        Fonts(PdfPageSelection.Parse(pageRanges), inspectionOptions, loadOptions);

    /// <summary>Inspects unique fonts declared by pages resolved by a document-relative selector.</summary>
    public PdfFontInventory Fonts(
        PdfPageSelector selector,
        PdfFontInspectionOptions? inspectionOptions = null,
        PdfLoadOptions? loadOptions = null) {
        Guard.NotNull(selector, nameof(selector));
        PdfReadDocument document = _document.GetReadDocument(loadOptions ?? _document.ReadOptions);
        if (document.Pages.Count == 0) throw new InvalidOperationException("PDF does not contain any readable pages.");
        return PdfFontInspector.Inspect(document, inspectionOptions, selector.ResolveSelection(document.Pages.Count));
    }

    /// <summary>Attempts to inspect fonts, returning preflight diagnostics when blocked or failed.</summary>
    public PdfOperationResult<PdfFontInventory> TryFonts(
        PdfFontInspectionOptions? inspectionOptions = null,
        PdfLoadOptions? loadOptions = null) =>
        _document.TryOperation(
            "Inspect fonts",
            PdfPreflightCapability.ReadLogicalObjects,
            () => Fonts(inspectionOptions, loadOptions),
            loadOptions ?? _document.ReadOptions);

    /// <summary>Builds a safe, immutable, bounded projection of the active PDF object graph.</summary>
    public PdfRawDocumentView RawStructure(
        PdfRawStructureOptions? structureOptions = null,
        PdfLoadOptions? loadOptions = null) =>
        _document.GetReadDocument(loadOptions ?? _document.ReadOptions).RawStructure(structureOptions);

    /// <summary>Attempts to build a bounded raw object view with preflight diagnostics.</summary>
    public PdfOperationResult<PdfRawDocumentView> TryRawStructure(
        PdfRawStructureOptions? structureOptions = null,
        PdfLoadOptions? loadOptions = null) =>
        _document.TryOperation(
            "Read raw structure",
            PdfPreflightCapability.ReadLogicalObjects,
            () => RawStructure(structureOptions, loadOptions),
            loadOptions ?? _document.ReadOptions);
}
