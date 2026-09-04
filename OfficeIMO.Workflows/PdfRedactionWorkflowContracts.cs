using OfficeIMO.Ocr;
using OfficeIMO.Pdf;

namespace OfficeIMO.Workflows;

/// <summary>Execution mode for a PDF redaction workflow.</summary>
public enum PdfRedactionWorkflowMode {
    /// <summary>Discover candidates and persist privacy-safe review data without modifying the PDF.</summary>
    PlanOnly,
    /// <summary>Re-plan the exact source, require reviewed decisions, apply approved candidates, and verify the output.</summary>
    ApplyAndVerify,
    /// <summary>Verify a previously produced output against approved source-bound geometry without rewriting it.</summary>
    VerifyExistingOutput
}

/// <summary>Controls when OCR participates in candidate discovery.</summary>
public enum PdfRedactionDetectionMode {
    /// <summary>Use the native PDF logical model only.</summary>
    NativeOnly,
    /// <summary>Use OCR only.</summary>
    OcrOnly,
    /// <summary>Use OCR only when native search finds no candidates.</summary>
    NativeThenOcr,
    /// <summary>Combine native and OCR candidates.</summary>
    NativeAndOcr
}

/// <summary>Policy for password-protected input PDFs.</summary>
public enum PdfRedactionEncryptedDocumentPolicy {
    /// <summary>Reject encrypted input even when credentials are supplied.</summary>
    Reject,
    /// <summary>Explicitly decrypt before redaction and publish an unencrypted result.</summary>
    Decrypt,
    /// <summary>Explicitly decrypt, redact, verify, and apply new Standard security.</summary>
    DecryptAndReencrypt
}

/// <summary>Explicit policy for signed sources and output derivatives.</summary>
public enum PdfRedactionSignaturePolicy {
    /// <summary>Reject a signed source and do not sign the output.</summary>
    RejectSignedSource,
    /// <summary>Remove invalidated signatures through an explicit full-rewrite derivative.</summary>
    CreateUnsignedDerivative,
    /// <summary>Create the derivative and append one caller-provided signature after redaction verification.</summary>
    CreateAndSignDerivative
}

/// <summary>Typed behavior for one serializable redaction search rule.</summary>
public enum PdfRedactionRuleKind {
    /// <summary>Literal text.</summary>
    Literal,
    /// <summary>Bounded regular expression.</summary>
    Regex,
    /// <summary>AcroForm field name.</summary>
    FormField,
    /// <summary><see cref="PdfLogicalElementKind"/> name.</summary>
    LogicalKind,
    /// <summary>Existing standard PDF /Redact annotations.</summary>
    RedactAnnotations
}

/// <summary>One serializable redaction search rule.</summary>
public sealed class PdfRedactionRule {
    /// <summary>Stable non-sensitive identifier included in review evidence. Use letters, digits, dot, underscore, or hyphen.</summary>
    public required string Name { get; set; }
    /// <summary>Typed rule behavior.</summary>
    public PdfRedactionRuleKind Kind { get; set; }
    /// <summary>Rule value. Required except for <see cref="PdfRedactionRuleKind.RedactAnnotations"/>; omitted from result evidence.</summary>
    public string? Value { get; set; }
    /// <summary>Intersecting content removal policy for candidates produced by this rule.</summary>
    public PdfRedactionContentScope ContentScope { get; set; } = PdfRedactionContentScope.TextAndUnderlay;
    /// <summary>Visible privacy-appearance policy for candidates produced by this rule.</summary>
    public PdfRedactionAppearanceMode AppearanceMode { get; set; } = PdfRedactionAppearanceMode.Exact;
}

/// <summary>Serializable point used by recipe geometry.</summary>
public sealed class PdfRedactionRecipePoint {
    /// <summary>Horizontal PDF user-space coordinate.</summary>
    public double X { get; set; }
    /// <summary>Vertical PDF user-space coordinate.</summary>
    public double Y { get; set; }
}

/// <summary>Serializable review geometry in a reusable recipe.</summary>
public sealed class PdfRedactionRecipeRegion {
    /// <summary>Stable non-sensitive identifier included in review evidence. Use letters, digits, dot, underscore, or hyphen.</summary>
    public required string Name { get; set; }
    /// <summary>Typed review geometry.</summary>
    public PdfRedactionRegionKind Kind { get; set; } = PdfRedactionRegionKind.Rectangle;
    /// <summary>One-based page number.</summary>
    public int PageNumber { get; set; } = 1;
    /// <summary>Rectangle left coordinate.</summary>
    public double X { get; set; }
    /// <summary>Rectangle bottom coordinate.</summary>
    public double Y { get; set; }
    /// <summary>Rectangle width.</summary>
    public double Width { get; set; }
    /// <summary>Rectangle height.</summary>
    public double Height { get; set; }
    /// <summary>Path points for quadrilateral, polygon, or freehand geometry.</summary>
    public IList<PdfRedactionRecipePoint> Points { get; set; } = new List<PdfRedactionRecipePoint>();
    /// <summary>Grouped rectangles.</summary>
    public IList<PdfRedactionRecipeRegion> Areas { get; set; } = new List<PdfRedactionRecipeRegion>();
    /// <summary>Freehand stroke width.</summary>
    public double StrokeWidth { get; set; } = 6D;
    /// <summary>Optional review label. Do not place sensitive matched text in labels.</summary>
    public string? Label { get; set; }
    /// <summary>Intersecting content removal policy for this explicit region.</summary>
    public PdfRedactionContentScope ContentScope { get; set; } = PdfRedactionContentScope.TextAndUnderlay;
    /// <summary>Visible privacy-appearance policy for this explicit region.</summary>
    public PdfRedactionAppearanceMode AppearanceMode { get; set; } = PdfRedactionAppearanceMode.Exact;
}

/// <summary>Versioned, reusable PDF redaction recipe.</summary>
public sealed class PdfRedactionRecipe {
    /// <summary>Stable recipe schema.</summary>
    public const string CurrentSchema = "officeimo.pdf.redaction.recipe.v1";
    /// <summary>Schema identifier.</summary>
    public string Schema { get; set; } = CurrentSchema;
    /// <summary>Native/OCR discovery policy.</summary>
    public PdfRedactionDetectionMode DetectionMode { get; set; } = PdfRedactionDetectionMode.NativeOnly;
    /// <summary>Case-sensitive literal matching when true.</summary>
    public bool MatchCase { get; set; }
    /// <summary>Per-match regex timeout in milliseconds.</summary>
    public int RegexTimeoutMilliseconds { get; set; } = 2_000;
    /// <summary>Search rules.</summary>
    public IList<PdfRedactionRule> Rules { get; set; } = new List<PdfRedactionRule>();
    /// <summary>Explicit review regions.</summary>
    public IList<PdfRedactionRecipeRegion> Regions { get; set; } = new List<PdfRedactionRecipeRegion>();
    /// <summary>Cleanup policy applied during destructive redaction.</summary>
    public PdfRedactionCleanupScope CleanupScope { get; set; } = PdfRedactionCleanupScope.All;
    /// <summary>Whether intersecting page-level vector paths are removed.</summary>
    public bool RemoveIntersectingPaths { get; set; } = true;
    /// <summary>Policy for unsupported image rewrites.</summary>
    public PdfRedactionUnsupportedImagePolicy UnsupportedImagePolicy { get; set; } = PdfRedactionUnsupportedImagePolicy.FailClosed;
    /// <summary>Policy for encrypted input.</summary>
    public PdfRedactionEncryptedDocumentPolicy EncryptedDocumentPolicy { get; set; } = PdfRedactionEncryptedDocumentPolicy.Reject;
    /// <summary>Signed-source and output-derivative policy.</summary>
    public PdfRedactionSignaturePolicy SignaturePolicy { get; set; } = PdfRedactionSignaturePolicy.RejectSignedSource;
}

/// <summary>Redaction-specific resource limits.</summary>
public sealed class PdfRedactionWorkflowLimits {
    /// <summary>Maximum input bytes.</summary>
    public long MaximumInputBytes { get; set; } = 256L * 1024L * 1024L;
    /// <summary>Maximum output bytes.</summary>
    public long MaximumOutputBytes { get; set; } = 512L * 1024L * 1024L;
    /// <summary>Maximum serialized privacy-safe evidence bytes for one item.</summary>
    public long MaximumEvidenceBytes { get; set; } = 64L * 1024L * 1024L;
    /// <summary>Maximum aggregate output and evidence bytes retained while preparing one atomic batch.</summary>
    public long MaximumBatchPreparedBytes { get; set; } = 1L * 1024L * 1024L * 1024L;
    /// <summary>Maximum recipe rules.</summary>
    public int MaximumRules { get; set; } = 500;
    /// <summary>Maximum aggregate characters in rule kinds and values.</summary>
    public int MaximumRuleCharacters { get; set; } = 1_000_000;
    /// <summary>Maximum normalized redaction areas.</summary>
    public int MaximumAreas { get; set; } = 10_000;
    /// <summary>Maximum aggregate geometry points and nested area declarations.</summary>
    public int MaximumGeometryPoints { get; set; } = 100_000;
    /// <summary>Maximum discovered candidates.</summary>
    public int MaximumCandidates { get; set; } = 25_000;
    /// <summary>Maximum items in one atomic batch.</summary>
    public int MaximumBatchItems { get; set; } = 100;
    /// <summary>Maximum items prepared concurrently before transactional publication.</summary>
    public int MaximumConcurrency { get; set; } = 1;
}

/// <summary>One source-bound reviewed decision manifest.</summary>
public sealed class PdfRedactionDecisionManifest {
    /// <summary>Stable manifest schema.</summary>
    public const string CurrentSchema = "officeimo.pdf.redaction.decisions.v1";
    /// <summary>Schema identifier.</summary>
    public string Schema { get; set; } = CurrentSchema;
    /// <summary>SHA-256 of exact source bytes.</summary>
    public required string SourceSha256 { get; set; }
    /// <summary>SHA-256 of the canonical recipe JSON.</summary>
    public required string RecipeSha256 { get; set; }
    /// <summary>Candidate identifiers approved for destructive application.</summary>
    public IList<string> ApprovedCandidateIds { get; set; } = new List<string>();
    /// <summary>Candidate identifiers explicitly rejected during review.</summary>
    public IList<string> RejectedCandidateIds { get; set; } = new List<string>();
}

/// <summary>One privacy-safe redaction candidate.</summary>
public sealed class PdfRedactionWorkflowArea {
    internal PdfRedactionWorkflowArea(PdfRedactionArea area) {
        PageNumber = area.PageNumber; X = area.X; Y = area.Y; Width = area.Width; Height = area.Height;
        Kind = area.ExactGeometry?.Kind ?? PdfRedactionRegionKind.Rectangle;
        Points = area.ExactGeometry?.Points.Select(static point => new PdfRedactionRecipePoint { X = point.X, Y = point.Y }).ToArray()
            ?? Array.Empty<PdfRedactionRecipePoint>();
        StrokeWidth = area.ExactGeometry?.StrokeWidth ?? 0D;
    }
    /// <summary>One-based page number.</summary>
    public int PageNumber { get; }
    /// <summary>Left coordinate.</summary>
    public double X { get; }
    /// <summary>Bottom coordinate.</summary>
    public double Y { get; }
    /// <summary>Width.</summary>
    public double Width { get; }
    /// <summary>Height.</summary>
    public double Height { get; }
    /// <summary>Exact reviewed geometry kind.</summary>
    public PdfRedactionRegionKind Kind { get; }
    /// <summary>Exact path points for polygon, quadrilateral, or freehand geometry.</summary>
    public IReadOnlyList<PdfRedactionRecipePoint> Points { get; }
    /// <summary>Exact freehand stroke width, or zero for other geometry.</summary>
    public double StrokeWidth { get; }
}

/// <summary>One privacy-safe, atomically reviewable redaction candidate.</summary>
public sealed class PdfRedactionWorkflowCandidate {
    internal PdfRedactionWorkflowCandidate(string id, string origin, string ruleName, PdfRedactionContentScope contentScope, PdfRedactionAppearanceMode appearanceMode, IReadOnlyList<PdfRedactionArea> areas, double? confidence, string? provider, string? model, string? language) {
        if (areas.Count == 0) throw new ArgumentException("A candidate requires at least one area.", nameof(areas));
        Id = id; Origin = origin; RuleName = ruleName; ContentScope = contentScope; AppearanceMode = appearanceMode; Areas = areas.Select(static area => new PdfRedactionWorkflowArea(area)).ToArray();
        PageNumber = areas[0].PageNumber;
        X = areas.Min(static area => area.X); Y = areas.Min(static area => area.Y);
        double right = areas.Max(static area => area.Right); double top = areas.Max(static area => area.Top);
        Width = right - X; Height = top - Y;
        Confidence = confidence; Provider = provider; Model = model; Language = language;
    }
    /// <summary>Source- and geometry-bound stable candidate identifier.</summary>
    public string Id { get; }
    /// <summary>Native, OCR, region, or annotation origin.</summary>
    public string Origin { get; }
    /// <summary>Stable rule or explicit-region name without matched content.</summary>
    public string RuleName { get; }
    /// <summary>Reviewed intersecting-content removal policy.</summary>
    public PdfRedactionContentScope ContentScope { get; }
    /// <summary>Reviewed visible privacy-appearance policy.</summary>
    public PdfRedactionAppearanceMode AppearanceMode { get; }
    /// <summary>One-based page number.</summary>
    public int PageNumber { get; }
    /// <summary>Left coordinate.</summary>
    public double X { get; }
    /// <summary>Bottom coordinate.</summary>
    public double Y { get; }
    /// <summary>Width.</summary>
    public double Width { get; }
    /// <summary>Height.</summary>
    public double Height { get; }
    /// <summary>Complete normalized area set governed by this one approval decision.</summary>
    public IReadOnlyList<PdfRedactionWorkflowArea> Areas { get; }
    /// <summary>Lowest OCR confidence for this candidate, when OCR-derived.</summary>
    public double? Confidence { get; }
    /// <summary>OCR provider identifier, when available.</summary>
    public string? Provider { get; }
    /// <summary>OCR model identifier, when available.</summary>
    public string? Model { get; }
    /// <summary>OCR language, when available.</summary>
    public string? Language { get; }
}

/// <summary>Runtime-only request for one redaction workflow.</summary>
public sealed class PdfRedactionWorkflowRequest {
    /// <summary>Caller request identifier.</summary>
    public string Id { get; set; } = Guid.NewGuid().ToString("N");
    /// <summary>Workflow mode.</summary>
    public PdfRedactionWorkflowMode Mode { get; set; }
    /// <summary>Input PDF path.</summary>
    public required string InputPath { get; set; }
    /// <summary>Output PDF path for apply, or existing output PDF path for verify.</summary>
    public string? OutputPath { get; set; }
    /// <summary>Optional privacy-safe JSON result/evidence path.</summary>
    public string? EvidencePath { get; set; }
    /// <summary>Additional runtime input paths, such as recipe and decision files, which publication must never replace.</summary>
    public IList<string> ProtectedInputPaths { get; set; } = new List<string>();
    /// <summary>Reusable recipe.</summary>
    public required PdfRedactionRecipe Recipe { get; set; }
    /// <summary>Required complete review decisions for apply and verify.</summary>
    public PdfRedactionDecisionManifest? Decisions { get; set; }
    /// <summary>Runtime OCR engine. It is never serialized into evidence.</summary>
    public IOcrEngine? OcrEngine { get; set; }
    /// <summary>OCR adapter limits and provider options. They are never copied wholesale into evidence.</summary>
    public OfficeIMO.Pdf.Ocr.PdfOcrMergeOptions? OcrOptions { get; set; }
    /// <summary>Runtime owner password for protected input. It is never serialized into evidence.</summary>
    public string? OwnerPassword { get; set; }
    /// <summary>Runtime new encryption for DecryptAndReencrypt. Passwords are never serialized into evidence.</summary>
    public PdfStandardEncryptionOptions? OutputEncryption { get; set; }
    /// <summary>Runtime signer used only by CreateAndSignDerivative. It is never serialized into evidence.</summary>
    public IPdfExternalSigner? OutputSigner { get; set; }
    /// <summary>Runtime external-signature settings. They are never serialized wholesale into evidence.</summary>
    public PdfExternalSignatureOptions? OutputSignatureOptions { get; set; }
    /// <summary>Optional runtime cryptographic validator for the derivative signature.</summary>
    public IPdfSignatureCryptographyProvider? OutputSignatureValidator { get; set; }
    /// <summary>Optional cancellation-aware independent validators applied to the final redacted artifact.</summary>
    public IList<IPdfRedactionCancellationAwareExternalValidator> ExternalValidators { get; set; } = new List<IPdfRedactionCancellationAwareExternalValidator>();
    /// <summary>Trusted expected output SHA-256 for zero-area verification when a security rewrite prevents byte comparison with the source.</summary>
    public string? ExpectedOutputSha256 { get; set; }
    /// <summary>Output conflict behavior.</summary>
    public OfficeWorkflowConflictPolicy ConflictPolicy { get; set; } = OfficeWorkflowConflictPolicy.Fail;
    /// <summary>Resource limits.</summary>
    public PdfRedactionWorkflowLimits Limits { get; set; } = new();
}

/// <summary>Runs reusable, source-bound PDF redaction workflows.</summary>
public interface IPdfRedactionWorkflowRunner {
    /// <summary>Plans, applies, or verifies one PDF redaction workflow.</summary>
    Task<PdfRedactionWorkflowResult> RunRedactionAsync(PdfRedactionWorkflowRequest request, IProgress<OfficeWorkflowProgress>? progress = null, CancellationToken cancellationToken = default);
    /// <summary>Runs a bounded all-or-nothing publication batch.</summary>
    Task<PdfRedactionBatchResult> RunRedactionBatchAsync(IEnumerable<PdfRedactionWorkflowRequest> requests, IProgress<OfficeWorkflowProgress>? progress = null, CancellationToken cancellationToken = default);
    /// <summary>Resolves and runs a deterministic file-set batch.</summary>
    Task<PdfRedactionBatchResult> RunRedactionBatchAsync(PdfRedactionBatchRequest request, IProgress<OfficeWorkflowProgress>? progress = null, CancellationToken cancellationToken = default);
}
