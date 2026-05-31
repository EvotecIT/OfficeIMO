namespace OfficeIMO.Pdf;

/// <summary>
/// Wrapper-friendly validation result for checking whether a PDF can be read or safely rewritten by OfficeIMO.Pdf.
/// </summary>
public sealed class PdfValidationResult {
    internal PdfValidationResult(PdfDocumentPreflight preflight) {
        Preflight = preflight;
    }

    /// <summary>Underlying preflight report with detailed capability and blocker information.</summary>
    public PdfDocumentPreflight Preflight { get; }

    /// <summary>True when OfficeIMO.Pdf can parse enough of the document for read-oriented operations.</summary>
    public bool IsValid => Preflight.CanRead;

    /// <summary>True when OfficeIMO.Pdf can parse enough of the document for read-oriented operations.</summary>
    public bool CanRead => Preflight.CanRead;

    /// <summary>True when OfficeIMO.Pdf can attempt rewrite-style operations without known blockers.</summary>
    public bool CanRewrite => Preflight.CanRewrite;

    /// <summary>True when OfficeIMO.Pdf can extract text from the document.</summary>
    public bool CanExtractText => Preflight.CanExtractText;

    /// <summary>True when OfficeIMO.Pdf can extract images from the document.</summary>
    public bool CanExtractImages => Preflight.CanExtractImages;

    /// <summary>True when OfficeIMO.Pdf can load logical readback objects from the document.</summary>
    public bool CanReadLogicalObjects => Preflight.CanReadLogicalObjects;

    /// <summary>True when OfficeIMO.Pdf can attempt page manipulation helpers without known blockers.</summary>
    public bool CanManipulatePages => Preflight.CanManipulatePages;

    /// <summary>True when OfficeIMO.Pdf can fill supported simple AcroForm fields.</summary>
    public bool CanFillSimpleFormFields => Preflight.CanFillSimpleFormFields;

    /// <summary>True when OfficeIMO.Pdf can flatten supported simple AcroForm fields.</summary>
    public bool CanFlattenSimpleFormFields => Preflight.CanFlattenSimpleFormFields;

    /// <summary>True when OfficeIMO.Pdf can fill and flatten supported simple AcroForm fields.</summary>
    public bool CanFillAndFlattenSimpleFormFields => Preflight.CanFillAndFlattenSimpleFormFields;

    /// <summary>Parsed document information when validation reached inspection.</summary>
    public PdfDocumentInfo? DocumentInfo => Preflight.DocumentInfo;

    /// <summary>Lightweight PDF markers read before full parsing.</summary>
    public PdfDocumentProbe Probe => Preflight.Probe;

    /// <summary>PDF header version when one was discovered.</summary>
    public string? HeaderVersion => Preflight.Probe.HeaderVersion;

    /// <summary>Page count when the document could be inspected; otherwise 0.</summary>
    public int PageCount => Preflight.DocumentInfo?.PageCount ?? 0;

    /// <summary>Human-readable diagnostics explaining validation failures or rewrite blockers.</summary>
    public IReadOnlyList<string> Diagnostics => Preflight.Diagnostics;

    /// <summary>Structured reasons why read-oriented validation failed.</summary>
    public IReadOnlyList<PdfReadBlocker> ReadBlockers => Preflight.ReadBlockers;

    /// <summary>Structured reasons why rewrite-style operations are blocked.</summary>
    public IReadOnlyList<PdfRewriteBlocker> RewriteBlockers => Preflight.RewriteBlockers;

    /// <summary>Returns true when a specific read blocker is present.</summary>
    public bool HasReadBlocker(PdfReadBlockerKind kind) => Preflight.HasReadBlocker(kind);

    /// <summary>Returns true when a specific rewrite blocker is present.</summary>
    public bool HasRewriteBlocker(PdfRewriteBlockerKind kind) => Preflight.HasRewriteBlocker(kind);

    /// <summary>Returns true when a specific read, extraction, or manipulation capability is available.</summary>
    public bool Can(PdfPreflightCapability capability) => Preflight.Can(capability);

    /// <summary>Returns diagnostics explaining why a specific capability is unavailable.</summary>
    public IReadOnlyList<string> GetCapabilityDiagnostics(PdfPreflightCapability capability) => Preflight.GetCapabilityDiagnostics(capability);
}
