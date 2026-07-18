namespace OfficeIMO.Pdf;

public sealed partial class PdfDocument {
    /// <summary>Compares this document with another PDF through the managed renderer and returns review artifacts.</summary>
    public PdfVisualComparisonReport CompareVisual(
        byte[] actualPdf,
        PdfPageSelection? selection = null,
        PdfVisualComparisonOptions? options = null,
        PdfReadOptions? actualReadOptions = null) {
        return PdfVisualComparer.Compare(GetBytesForOperation(), actualPdf, selection, options, ReadOptions, actualReadOptions);
    }

    /// <summary>Compares this document with another fluent PDF through the managed renderer.</summary>
    public PdfVisualComparisonReport CompareVisual(
        PdfDocument actualDocument,
        PdfPageSelection? selection = null,
        PdfVisualComparisonOptions? options = null) {
        Guard.NotNull(actualDocument, nameof(actualDocument));
        return PdfVisualComparer.Compare(GetBytesForOperation(), actualDocument.GetBytesForOperation(), selection, options, ReadOptions, actualDocument.ReadOptions);
    }
    /// <summary>
    /// Compares this PDF with a rewritten PDF and reports whether important document signals were preserved.
    /// </summary>
    public PdfRewritePreservationReport AssessRewritePreservation(PdfDocument rewrittenDocument, PdfRewritePreservationOptions? options = null) {
        Guard.NotNull(rewrittenDocument, nameof(rewrittenDocument));
        return AssessRewritePreservation(rewrittenDocument.GetBytesForOperation(), options);
    }

    /// <summary>
    /// Compares this PDF with rewritten PDF bytes and reports whether important document signals were preserved.
    /// </summary>
    public PdfRewritePreservationReport AssessRewritePreservation(byte[] rewrittenPdf, PdfRewritePreservationOptions? options = null) {
        return PdfRewritePreservation.Assess(GetBytesForOperation(), rewrittenPdf, options);
    }

    /// <summary>
    /// Compares this PDF with a rewritten PDF stream and reports whether important document signals were preserved.
    /// </summary>
    public PdfRewritePreservationReport AssessRewritePreservation(Stream rewrittenStream, PdfRewritePreservationOptions? options = null) {
        return AssessRewritePreservation(ReadProofStream(rewrittenStream), options);
    }

    /// <summary>
    /// Compares this PDF with a rewritten PDF file and reports whether important document signals were preserved.
    /// </summary>
    public PdfRewritePreservationReport AssessRewritePreservation(string rewrittenPath, PdfRewritePreservationOptions? options = null) {
        Guard.NotNullOrWhiteSpace(rewrittenPath, nameof(rewrittenPath));
        return AssessRewritePreservation(File.ReadAllBytes(rewrittenPath), options);
    }

    /// <summary>
    /// Compares this PDF with a rewritten PDF and throws when important document signals were not preserved.
    /// </summary>
    public PdfRewritePreservationReport AssertRewritePreserved(PdfDocument rewrittenDocument, PdfRewritePreservationOptions? options = null) {
        Guard.NotNull(rewrittenDocument, nameof(rewrittenDocument));
        return AssertRewritePreserved(rewrittenDocument.GetBytesForOperation(), options);
    }

    /// <summary>
    /// Compares this PDF with rewritten PDF bytes and throws when important document signals were not preserved.
    /// </summary>
    public PdfRewritePreservationReport AssertRewritePreserved(byte[] rewrittenPdf, PdfRewritePreservationOptions? options = null) {
        return PdfRewritePreservation.AssertPreserved(GetBytesForOperation(), rewrittenPdf, options);
    }

    /// <summary>
    /// Compares this PDF with a rewritten PDF stream and throws when important document signals were not preserved.
    /// </summary>
    public PdfRewritePreservationReport AssertRewritePreserved(Stream rewrittenStream, PdfRewritePreservationOptions? options = null) {
        return AssertRewritePreserved(ReadProofStream(rewrittenStream), options);
    }

    /// <summary>
    /// Compares this PDF with a rewritten PDF file and throws when important document signals were not preserved.
    /// </summary>
    public PdfRewritePreservationReport AssertRewritePreserved(string rewrittenPath, PdfRewritePreservationOptions? options = null) {
        Guard.NotNullOrWhiteSpace(rewrittenPath, nameof(rewrittenPath));
        return AssertRewritePreserved(File.ReadAllBytes(rewrittenPath), options);
    }

    /// <summary>
    /// Runs a one-row rewrite-preservation matrix scenario from this PDF using a normal PdfDocument rewrite operation.
    /// </summary>
    public PdfRewritePreservationMatrixReport AssessRewritePreservationMatrix(
        string id,
        string operation,
        Func<PdfDocument, PdfDocument> rewrite,
        PdfRewritePreservationOptions? options = null) {
        return AssessRewritePreservationMatrix(
            id,
            operation,
            rewrite,
            PdfRewritePreservationMatrixClassification.RewriteSafe,
            options,
            sourceFeatures: null);
    }

    /// <summary>
    /// Runs a one-row rewrite-preservation matrix scenario from this PDF with source feature labels.
    /// </summary>
    public PdfRewritePreservationMatrixReport AssessRewritePreservationMatrix(
        string id,
        string operation,
        Func<PdfDocument, PdfDocument> rewrite,
        PdfRewritePreservationOptions? options,
        IEnumerable<string>? sourceFeatures) {
        return AssessRewritePreservationMatrix(
            id,
            operation,
            rewrite,
            PdfRewritePreservationMatrixClassification.RewriteSafe,
            options,
            sourceFeatures);
    }

    /// <summary>
    /// Runs a one-row rewrite-preservation matrix scenario from this PDF using a normal PdfDocument rewrite operation and expected outcome.
    /// </summary>
    public PdfRewritePreservationMatrixReport AssessRewritePreservationMatrix(
        string id,
        string operation,
        Func<PdfDocument, PdfDocument> rewrite,
        PdfRewritePreservationMatrixClassification expectedClassification,
        PdfRewritePreservationOptions? options = null,
        IEnumerable<string>? sourceFeatures = null) {
        PdfRewritePreservationMatrixScenario scenario = CreateRewritePreservationMatrixScenario(id, operation, rewrite, expectedClassification, options, sourceFeatures);
        return PdfRewritePreservationMatrix.Assess(new[] { scenario });
    }

    /// <summary>
    /// Runs a one-row rewrite-preservation matrix scenario and throws when the observed outcome differs from the expected outcome.
    /// </summary>
    public PdfRewritePreservationMatrixReport AssertRewritePreservationMatrix(
        string id,
        string operation,
        Func<PdfDocument, PdfDocument> rewrite,
        PdfRewritePreservationOptions? options = null) {
        return AssertRewritePreservationMatrix(
            id,
            operation,
            rewrite,
            PdfRewritePreservationMatrixClassification.RewriteSafe,
            options,
            sourceFeatures: null);
    }

    /// <summary>
    /// Runs a one-row rewrite-preservation matrix scenario with source feature labels and throws when preservation failed.
    /// </summary>
    public PdfRewritePreservationMatrixReport AssertRewritePreservationMatrix(
        string id,
        string operation,
        Func<PdfDocument, PdfDocument> rewrite,
        PdfRewritePreservationOptions? options,
        IEnumerable<string>? sourceFeatures) {
        return AssertRewritePreservationMatrix(
            id,
            operation,
            rewrite,
            PdfRewritePreservationMatrixClassification.RewriteSafe,
            options,
            sourceFeatures);
    }

    /// <summary>
    /// Runs a one-row rewrite-preservation matrix scenario and throws when the observed outcome differs from the expected outcome.
    /// </summary>
    public PdfRewritePreservationMatrixReport AssertRewritePreservationMatrix(
        string id,
        string operation,
        Func<PdfDocument, PdfDocument> rewrite,
        PdfRewritePreservationMatrixClassification expectedClassification,
        PdfRewritePreservationOptions? options = null,
        IEnumerable<string>? sourceFeatures = null) {
        PdfRewritePreservationMatrixScenario scenario = CreateRewritePreservationMatrixScenario(id, operation, rewrite, expectedClassification, options, sourceFeatures);
        return PdfRewritePreservationMatrix.AssertExpected(new[] { scenario });
    }

    /// <summary>
    /// Verifies that configured redaction markers were removed and retained markers remain readable in this PDF.
    /// </summary>
    public PdfRedactionVerificationReport VerifyRedactions(PdfRedactionVerificationOptions options) {
        return PdfRedactionVerification.Verify(GetBytesForOperation(), options);
    }

    /// <summary>
    /// Verifies configured redaction markers and throws when removed content remains or retained content disappeared.
    /// </summary>
    public PdfRedactionVerificationReport AssertRedactionsVerified(PdfRedactionVerificationOptions options) {
        return PdfRedactionVerification.AssertVerified(GetBytesForOperation(), options);
    }

    private PdfRewritePreservationMatrixScenario CreateRewritePreservationMatrixScenario(
        string id,
        string operation,
        Func<PdfDocument, PdfDocument> rewrite,
        PdfRewritePreservationMatrixClassification expectedClassification,
        PdfRewritePreservationOptions? options,
        IEnumerable<string>? sourceFeatures) {
        Guard.NotNull(rewrite, nameof(rewrite));

        byte[] sourcePdf = GetBytesForOperation();
        var scenario = new PdfRewritePreservationMatrixScenario(
            id,
            operation,
            sourcePdf,
            pdf => rewrite(Open(pdf)).GetBytesForOperation()) {
                ExpectedClassification = expectedClassification,
                PreservationOptions = options
            };

        if (sourceFeatures is not null) {
            scenario.WithSourceFeatures(sourceFeatures.ToArray());
        }

        return scenario;
    }

    private static byte[] ReadProofStream(Stream stream) {
        Guard.NotNull(stream, nameof(stream));
        if (!stream.CanRead) {
            throw new ArgumentException("Stream must be readable.", nameof(stream));
        }

        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return buffer.ToArray();
    }
}
