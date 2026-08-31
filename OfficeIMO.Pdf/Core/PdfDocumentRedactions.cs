namespace OfficeIMO.Pdf;

/// <summary>Permanent redaction operations for one PDF document.</summary>
public sealed class PdfDocumentRedactions {
    private readonly PdfDocument _document;

    internal PdfDocumentRedactions(PdfDocument document) => _document = document;

    /// <summary>Plans rectangle-based redaction impact without modifying the PDF.</summary>
    public PdfRedactionPlan Plan(IEnumerable<PdfRedactionArea> areas, PdfTextLayoutOptions? layoutOptions = null, PdfReadOptions? options = null) =>
        _document.PlanRedactions(areas, layoutOptions, options);

    /// <summary>Searches literal text, regex, logical kinds, and form-field names into a reviewable plan.</summary>
    public PdfRedactionPlan Search(PdfRedactionSearchOptions search, PdfTextLayoutOptions? layoutOptions = null, PdfReadOptions? options = null) =>
        _document.SearchRedactions(search, layoutOptions, options);

    /// <summary>Creates a new PDF with content and annotations removed from the supplied areas.</summary>
    public PdfDocument Apply(IEnumerable<PdfRedactionArea> areas, PdfRedactionApplyOptions? applyOptions = null, PdfTextLayoutOptions? layoutOptions = null, PdfReadOptions? options = null) =>
        _document.ApplyRedactions(areas, applyOptions, layoutOptions, options);

    /// <summary>Applies a reviewed redaction plan, including exact field removal.</summary>
    public PdfDocument Apply(PdfRedactionPlan plan, PdfRedactionApplyOptions? applyOptions = null, PdfTextLayoutOptions? layoutOptions = null, PdfReadOptions? options = null) =>
        _document.ApplyRedactions(plan, applyOptions, layoutOptions, options);

    /// <summary>Attempts to apply redactions and returns preflight diagnostics when blocked.</summary>
    public PdfOperationResult<PdfDocument> TryApply(IEnumerable<PdfRedactionArea> areas, PdfRedactionApplyOptions? applyOptions = null, PdfTextLayoutOptions? layoutOptions = null, PdfReadOptions? options = null) =>
        _document.TryApplyRedactions(areas, applyOptions, layoutOptions, options);

    /// <summary>Verifies configured removed and retained markers in the rewritten PDF.</summary>
    public PdfRedactionVerificationReport Verify(PdfRedactionVerificationOptions options) => _document.VerifyRedactions(options);

    /// <summary>Verifies configured markers and throws when redaction proof fails.</summary>
    public PdfRedactionVerificationReport AssertVerified(PdfRedactionVerificationOptions options) => _document.AssertRedactionsVerified(options);

    /// <summary>Verifies configured markers and reports any content still intersecting the reviewed plan areas.</summary>
    public PdfRedactionVerificationReport VerifyAppliedPlan(PdfRedactionPlan reviewedPlan, PdfRedactionVerificationOptions options) =>
        PdfRedactionVerification.VerifyAppliedPlan(_document.ToBytes(), reviewedPlan, options, _document.ReadOptions);

    /// <summary>Verifies a reviewed plan and throws when planned content remains in its areas.</summary>
    public PdfRedactionVerificationReport AssertAppliedPlan(PdfRedactionPlan reviewedPlan, PdfRedactionVerificationOptions options) =>
        PdfRedactionVerification.AssertAppliedPlan(_document.ToBytes(), reviewedPlan, options, _document.ReadOptions);
}
