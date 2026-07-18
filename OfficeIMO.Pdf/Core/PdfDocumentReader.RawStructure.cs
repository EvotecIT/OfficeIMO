namespace OfficeIMO.Pdf;

public sealed partial class PdfDocumentReader {
    /// <summary>Builds a safe, immutable, bounded projection of the active raw object graph.</summary>
    public PdfRawDocumentView RawStructure(PdfRawStructureOptions? structureOptions = null, PdfReadOptions? readOptions = null) {
        return ReadDocument(readOptions).RawStructure(structureOptions);
    }

    /// <summary>Attempts to build a bounded raw object view with preflight diagnostics.</summary>
    public PdfOperationResult<PdfRawDocumentView> TryRawStructure(PdfRawStructureOptions? structureOptions = null, PdfReadOptions? options = null) {
        return _document.TryOperation(
            "Read raw structure",
            PdfPreflightCapability.ReadLogicalObjects,
            () => RawStructure(structureOptions, options),
            ResolveReadOptions(options));
    }
}
