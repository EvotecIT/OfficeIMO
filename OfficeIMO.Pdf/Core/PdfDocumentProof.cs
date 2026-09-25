namespace OfficeIMO.Pdf;

/// <summary>Visual and structural preservation proof operations for one PDF document.</summary>
public sealed class PdfDocumentProof {
    private readonly PdfDocument _document;

    internal PdfDocumentProof(PdfDocument document) => _document = document;

    /// <summary>Aligns pages with another PDF and reports insertion, deletion, reorder, and visual-change candidates.</summary>
    public PdfPageChangeReport AnalyzePageChanges(PdfDocument actualDocument, PdfPageChangeOptions? options = null, System.Threading.CancellationToken cancellationToken = default) {
        Guard.NotNull(actualDocument, nameof(actualDocument));
        return PdfPageChangeAnalyzer.Analyze(_document.GetBytesForOperation(cancellationToken), actualDocument.GetBytesForOperation(cancellationToken),
            options, _document.ReadOptions, actualDocument.ReadOptions, cancellationToken);
    }

    /// <summary>Renders and compares an aligned page pair from two documents.</summary>
    public PdfVisualPageComparison CompareVisualPages(int expectedPageNumber, PdfDocument actualDocument, int actualPageNumber,
        PdfVisualComparisonOptions? options = null, System.Threading.CancellationToken cancellationToken = default) {
        Guard.NotNull(actualDocument, nameof(actualDocument));
        return PdfVisualComparer.ComparePages(_document.GetBytesForOperation(cancellationToken), expectedPageNumber,
            actualDocument.GetBytesForOperation(cancellationToken), actualPageNumber, options,
            _document.ReadOptions, actualDocument.ReadOptions, cancellationToken);
    }

    /// <summary>Compares this document with PDF bytes through the managed renderer.</summary>
    public PdfVisualComparisonReport CompareVisual(byte[] actualPdf, PdfPageSelection? selection = null, PdfVisualComparisonOptions? options = null, PdfLoadOptions? actualReadOptions = null) =>
        _document.CompareVisual(actualPdf, selection, options, actualReadOptions);

    /// <summary>Compares this document with another PDF through the managed renderer.</summary>
    public PdfVisualComparisonReport CompareVisual(PdfDocument actualDocument, PdfPageSelection? selection = null, PdfVisualComparisonOptions? options = null) =>
        _document.CompareVisual(actualDocument, selection, options);

    /// <summary>Compares authenticated document snapshots with cooperative cancellation and bounded rendered artifacts.</summary>
    public PdfVisualComparisonReport CompareVisual(PdfDocument actualDocument, System.Threading.CancellationToken cancellationToken, PdfPageSelection? selection = null, PdfVisualComparisonOptions? options = null) =>
        _document.CompareVisual(actualDocument, cancellationToken, selection, options);

    /// <summary>Assesses structural preservation against a rewritten document.</summary>
    public PdfRewritePreservationReport AssessRewritePreservation(PdfDocument rewrittenDocument, PdfRewritePreservationOptions? options = null) =>
        _document.AssessRewritePreservation(rewrittenDocument, options);

    /// <summary>Assesses structural preservation against rewritten bytes.</summary>
    public PdfRewritePreservationReport AssessRewritePreservation(byte[] rewrittenPdf, PdfRewritePreservationOptions? options = null) =>
        _document.AssessRewritePreservation(rewrittenPdf, options);

    /// <summary>Assesses structural preservation against a rewritten stream.</summary>
    public PdfRewritePreservationReport AssessRewritePreservation(Stream rewrittenStream, PdfRewritePreservationOptions? options = null) =>
        _document.AssessRewritePreservation(rewrittenStream, options);

    /// <summary>Assesses structural preservation against a rewritten file.</summary>
    public PdfRewritePreservationReport AssessRewritePreservation(string rewrittenPath, PdfRewritePreservationOptions? options = null) =>
        _document.AssessRewritePreservation(rewrittenPath, options);

    /// <summary>Asserts structural preservation against a rewritten document.</summary>
    public PdfRewritePreservationReport AssertRewritePreserved(PdfDocument rewrittenDocument, PdfRewritePreservationOptions? options = null) =>
        _document.AssertRewritePreserved(rewrittenDocument, options);

    /// <summary>Asserts structural preservation against rewritten bytes.</summary>
    public PdfRewritePreservationReport AssertRewritePreserved(byte[] rewrittenPdf, PdfRewritePreservationOptions? options = null) =>
        _document.AssertRewritePreserved(rewrittenPdf, options);

    /// <summary>Asserts structural preservation against a rewritten stream.</summary>
    public PdfRewritePreservationReport AssertRewritePreserved(Stream rewrittenStream, PdfRewritePreservationOptions? options = null) =>
        _document.AssertRewritePreserved(rewrittenStream, options);

    /// <summary>Asserts structural preservation against a rewritten file.</summary>
    public PdfRewritePreservationReport AssertRewritePreserved(string rewrittenPath, PdfRewritePreservationOptions? options = null) =>
        _document.AssertRewritePreserved(rewrittenPath, options);

    /// <summary>Assesses one rewrite scenario using the default rewrite-safe expectation.</summary>
    public PdfRewritePreservationMatrixReport AssessRewritePreservationMatrix(string id, string operation, Func<PdfDocument, PdfDocument> rewrite, PdfRewritePreservationOptions? options = null) =>
        _document.AssessRewritePreservationMatrix(id, operation, rewrite, options);

    /// <summary>Assesses one rewrite scenario with source feature labels.</summary>
    public PdfRewritePreservationMatrixReport AssessRewritePreservationMatrix(string id, string operation, Func<PdfDocument, PdfDocument> rewrite, PdfRewritePreservationOptions? options, IEnumerable<string>? sourceFeatures) =>
        _document.AssessRewritePreservationMatrix(id, operation, rewrite, options, sourceFeatures);

    /// <summary>Assesses one rewrite scenario against an explicit expected classification.</summary>
    public PdfRewritePreservationMatrixReport AssessRewritePreservationMatrix(string id, string operation, Func<PdfDocument, PdfDocument> rewrite, PdfRewritePreservationMatrixClassification expectedClassification, PdfRewritePreservationOptions? options = null, IEnumerable<string>? sourceFeatures = null) =>
        _document.AssessRewritePreservationMatrix(id, operation, rewrite, expectedClassification, options, sourceFeatures);

    /// <summary>Asserts one rewrite scenario using the default rewrite-safe expectation.</summary>
    public PdfRewritePreservationMatrixReport AssertRewritePreservationMatrix(string id, string operation, Func<PdfDocument, PdfDocument> rewrite, PdfRewritePreservationOptions? options = null) =>
        _document.AssertRewritePreservationMatrix(id, operation, rewrite, options);

    /// <summary>Asserts one rewrite scenario with source feature labels.</summary>
    public PdfRewritePreservationMatrixReport AssertRewritePreservationMatrix(string id, string operation, Func<PdfDocument, PdfDocument> rewrite, PdfRewritePreservationOptions? options, IEnumerable<string>? sourceFeatures) =>
        _document.AssertRewritePreservationMatrix(id, operation, rewrite, options, sourceFeatures);

    /// <summary>Asserts one rewrite scenario against an explicit expected classification.</summary>
    public PdfRewritePreservationMatrixReport AssertRewritePreservationMatrix(string id, string operation, Func<PdfDocument, PdfDocument> rewrite, PdfRewritePreservationMatrixClassification expectedClassification, PdfRewritePreservationOptions? options = null, IEnumerable<string>? sourceFeatures = null) =>
        _document.AssertRewritePreservationMatrix(id, operation, rewrite, expectedClassification, options, sourceFeatures);
}
