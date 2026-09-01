namespace OfficeIMO.Pdf;

/// <summary>Lossless optimization operations for one PDF document.</summary>
public sealed class PdfDocumentOptimization {
    private readonly PdfDocument _document;

    internal PdfDocumentOptimization(PdfDocument document) => _document = document;

    /// <summary>Builds an optimization opportunity report without modifying the PDF.</summary>
    public PdfOptimizationReport Analyze(PdfLoadOptions? options = null) => _document.AnalyzeOptimization(options);

    /// <summary>Applies dependency-free lossless optimization.</summary>
    public PdfOptimizationActionResult Apply(PdfOptimizationOptions? options = null) => _document.Optimize(options);

    /// <summary>Applies a named deterministic lossless optimization profile.</summary>
    public PdfOptimizationActionResult Apply(PdfOptimizationProfile profile) => _document.Optimize(profile);
}
