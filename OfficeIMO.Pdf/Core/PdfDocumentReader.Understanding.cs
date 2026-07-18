namespace OfficeIMO.Pdf;

public sealed partial class PdfDocumentReader {
    /// <summary>Runs the pluggable text-understanding pipeline for all or selected pages.</summary>
    public PdfUnderstandingResult Understand(PdfUnderstandingPipelineOptions? options = null, PdfPageSelection? selection = null, PdfReadOptions? readOptions = null) =>
        new PdfUnderstandingPipeline(options).Run(ReadDocument(readOptions), selection);
}
