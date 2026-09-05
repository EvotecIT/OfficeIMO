namespace OfficeIMO.Pdf;

/// <summary>Fluent font-resource inspection operations for a <see cref="PdfDocument"/>.</summary>
internal sealed partial class PdfDocumentReader {
    /// <summary>Inspects unique fonts declared by all pages and nested Form XObjects.</summary>
    public PdfFontInventory Fonts(
        PdfFontInspectionOptions? inspectionOptions = null,
        PdfLoadOptions? readOptions = null) {
        return PdfFontInspector.Inspect(ReadDocument(readOptions), inspectionOptions);
    }

    /// <summary>Attempts to inspect fonts, returning preflight diagnostics when blocked or failed.</summary>
    public PdfOperationResult<PdfFontInventory> FontsResult(
        PdfFontInspectionOptions? inspectionOptions = null,
        PdfLoadOptions? readOptions = null) {
        return _document.TryOperation(
            "Inspect fonts",
            PdfPreflightCapability.ReadLogicalObjects,
            () => Fonts(inspectionOptions, readOptions),
            ResolveReadOptions(readOptions));
    }

    /// <summary>Inspects unique fonts declared by selected pages and nested Form XObjects.</summary>
    public PdfFontInventory Fonts(
        PdfPageSelection selection,
        PdfFontInspectionOptions? inspectionOptions = null,
        PdfLoadOptions? readOptions = null) {
        Guard.NotNull(selection, nameof(selection));
        return PdfFontInspector.Inspect(ReadDocument(readOptions), inspectionOptions, selection);
    }

    /// <summary>Attempts to inspect fonts on selected pages, returning preflight diagnostics when blocked or failed.</summary>
    public PdfOperationResult<PdfFontInventory> FontsResult(
        PdfPageSelection selection,
        PdfFontInspectionOptions? inspectionOptions = null,
        PdfLoadOptions? readOptions = null) {
        Guard.NotNull(selection, nameof(selection));
        return _document.TryOperation(
            "Inspect fonts",
            PdfPreflightCapability.ReadLogicalObjects,
            () => Fonts(selection, inspectionOptions, readOptions),
            ResolveReadOptions(readOptions));
    }

    /// <summary>Inspects unique fonts declared by pages described by one-based page ranges.</summary>
    public PdfFontInventory Fonts(
        string pageRanges,
        PdfFontInspectionOptions? inspectionOptions = null,
        PdfLoadOptions? readOptions = null) {
        return Fonts(PdfPageSelection.Parse(pageRanges), inspectionOptions, readOptions);
    }

    /// <summary>Attempts to inspect fonts on pages described by one-based page ranges.</summary>
    public PdfOperationResult<PdfFontInventory> FontsResult(
        string pageRanges,
        PdfFontInspectionOptions? inspectionOptions = null,
        PdfLoadOptions? readOptions = null) {
        return FontsResult(PdfPageSelection.Parse(pageRanges), inspectionOptions, readOptions);
    }
}
