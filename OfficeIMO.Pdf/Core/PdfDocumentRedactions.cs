namespace OfficeIMO.Pdf;

/// <summary>Permanent redaction operations for one PDF document.</summary>
public sealed class PdfDocumentRedactions {
    private readonly PdfDocument _document;

    internal PdfDocumentRedactions(PdfDocument document) => _document = document;

    /// <summary>Plans rectangle-based redaction impact without modifying the PDF.</summary>
    public PdfRedactionPlan Plan(IEnumerable<PdfRedactionArea> areas, PdfTextLayoutOptions? layoutOptions = null, PdfLoadOptions? options = null) =>
        _document.PlanRedactions(areas, layoutOptions, options);

    /// <summary>Searches literal text, regex, logical kinds, and form-field names into a reviewable plan.</summary>
    public PdfRedactionPlan Search(PdfRedactionSearchOptions search, PdfTextLayoutOptions? layoutOptions = null, PdfLoadOptions? options = null) =>
        _document.SearchRedactions(search, layoutOptions, options);

    /// <summary>Creates a new PDF with intersecting content and annotations removed from the supplied areas.</summary>
    /// <remarks>Text-show operations are rewritten at glyph granularity when the encoded mapping is safe; otherwise the complete PDF text object is removed.</remarks>
    public PdfDocument Apply(IEnumerable<PdfRedactionArea> areas, PdfRedactionApplyOptions? applyOptions = null, PdfTextLayoutOptions? layoutOptions = null, PdfLoadOptions? options = null) =>
        _document.ApplyRedactions(areas, applyOptions, layoutOptions, options);

    /// <summary>Applies a reviewed redaction plan, including exact field removal.</summary>
    public PdfDocument Apply(PdfRedactionPlan plan, PdfRedactionApplyOptions? applyOptions = null, PdfTextLayoutOptions? layoutOptions = null, PdfLoadOptions? options = null) =>
        _document.ApplyRedactions(plan, applyOptions, layoutOptions, options);

    /// <summary>
    /// Applies a source-bound reviewed plan and returns the rewritten PDF, selected mutation path,
    /// and actual-versus-planned evidence from the rewritten artifact.
    /// </summary>
    /// <remarks>When verification options are omitted, complete stream inspection and managed rendering checks are required.</remarks>
    public PdfRedactionApplyResult ApplyWithEvidence(
        PdfRedactionPlan plan,
        PdfRedactionApplyOptions? applyOptions = null,
        PdfRedactionVerificationOptions? verificationOptions = null,
        PdfTextLayoutOptions? layoutOptions = null,
        PdfLoadOptions? options = null) =>
        _document.ApplyRedactionsWithEvidence(plan, applyOptions, verificationOptions, layoutOptions, options);

    /// <summary>Attempts to apply redactions and returns preflight diagnostics when blocked.</summary>
    public PdfOperationResult<PdfDocument> TryApply(IEnumerable<PdfRedactionArea> areas, PdfRedactionApplyOptions? applyOptions = null, PdfTextLayoutOptions? layoutOptions = null, PdfLoadOptions? options = null) =>
        _document.TryApplyRedactions(areas, applyOptions, layoutOptions, options);

    /// <summary>Attempts to apply a reviewed plan and returns diagnostics instead of throwing when the mutation is blocked or fails.</summary>
    public PdfOperationResult<PdfRedactionApplyResult> TryApplyWithEvidence(
        PdfRedactionPlan plan,
        PdfRedactionApplyOptions? applyOptions = null,
        PdfRedactionVerificationOptions? verificationOptions = null,
        PdfTextLayoutOptions? layoutOptions = null,
        PdfLoadOptions? options = null) =>
        _document.TryApplyRedactionsWithEvidence(plan, applyOptions, verificationOptions, layoutOptions, options);

    /// <summary>Verifies configured removed and retained markers in the rewritten PDF.</summary>
    public PdfRedactionVerificationReport Verify(PdfRedactionVerificationOptions options) => _document.VerifyRedactions(options);

    /// <summary>Verifies configured markers and throws when redaction proof fails.</summary>
    public PdfRedactionVerificationReport AssertVerified(PdfRedactionVerificationOptions options) => _document.AssertRedactionsVerified(options);

    /// <summary>Verifies configured markers and reports any content still intersecting the reviewed plan areas.</summary>
    /// <remarks>The plan is source-bound when applied; this residual check does not independently prove rewrite lineage.</remarks>
    public PdfRedactionVerificationReport VerifyAppliedPlan(PdfRedactionPlan reviewedPlan, PdfRedactionVerificationOptions options) =>
        PdfRedactionVerification.VerifyAppliedPlan(_document.ToBytes(), reviewedPlan, options, _document.ReadOptions);

    /// <summary>Verifies a reviewed plan and throws when planned content remains in its areas.</summary>
    /// <remarks>The plan is source-bound when applied; this residual check does not independently prove rewrite lineage.</remarks>
    public PdfRedactionVerificationReport AssertAppliedPlan(PdfRedactionPlan reviewedPlan, PdfRedactionVerificationOptions options) =>
        PdfRedactionVerification.AssertAppliedPlan(_document.ToBytes(), reviewedPlan, options, _document.ReadOptions);
}
