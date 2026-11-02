namespace OfficeIMO.Pdf;

public sealed partial class PdfDocument {
    /// <summary>Builds an immutable, bounded view of the document's active raw object graph.</summary>
    public PdfRawDocumentView RawStructure(PdfRawStructureOptions? options = null, PdfLoadOptions? readOptions = null) =>
        Reader.RawStructure(options, readOptions);
}
