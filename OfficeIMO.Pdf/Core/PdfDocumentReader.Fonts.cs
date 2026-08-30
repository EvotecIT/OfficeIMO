namespace OfficeIMO.Pdf;

/// <summary>Fluent font-resource inspection operations for a <see cref="PdfDocument"/>.</summary>
public sealed partial class PdfDocumentReader {
    /// <summary>Inspects unique fonts declared by all pages and nested Form XObjects.</summary>
    public PdfFontInventory Fonts(
        PdfFontInspectionOptions? inspectionOptions = null,
        PdfReadOptions? readOptions = null) {
        return PdfFontInspector.Inspect(ReadDocument(readOptions), inspectionOptions);
    }

    /// <summary>Attempts to inspect fonts, returning preflight diagnostics when blocked or failed.</summary>
    public PdfOperationResult<PdfFontInventory> TryFonts(
        PdfFontInspectionOptions? inspectionOptions = null,
        PdfReadOptions? readOptions = null) {
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
        PdfReadOptions? readOptions = null) {
        Guard.NotNull(selection, nameof(selection));
        return PdfFontInspector.Inspect(ReadDocument(readOptions), inspectionOptions, selection);
    }

    /// <summary>Attempts to inspect fonts on selected pages, returning preflight diagnostics when blocked or failed.</summary>
    public PdfOperationResult<PdfFontInventory> TryFonts(
        PdfPageSelection selection,
        PdfFontInspectionOptions? inspectionOptions = null,
        PdfReadOptions? readOptions = null) {
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
        PdfReadOptions? readOptions = null) {
        return Fonts(PdfPageSelection.Parse(pageRanges), inspectionOptions, readOptions);
    }

    /// <summary>Attempts to inspect fonts on pages described by one-based page ranges.</summary>
    public PdfOperationResult<PdfFontInventory> TryFonts(
        string pageRanges,
        PdfFontInspectionOptions? inspectionOptions = null,
        PdfReadOptions? readOptions = null) {
        return TryFonts(PdfPageSelection.Parse(pageRanges), inspectionOptions, readOptions);
    }
}
