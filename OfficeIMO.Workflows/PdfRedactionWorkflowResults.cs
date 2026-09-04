namespace OfficeIMO.Workflows;

/// <summary>Privacy-safe verification issue summary.</summary>
public sealed class PdfRedactionEvidenceIssue {
    internal PdfRedactionEvidenceIssue(string feature) => Feature = feature;
    /// <summary>Stable verification feature code. Sensitive marker values and messages are omitted.</summary>
    public string Feature { get; }
}

/// <summary>Privacy-safe evidence for one redaction workflow.</summary>
public sealed class PdfRedactionWorkflowEvidence {
    internal PdfRedactionWorkflowEvidence(string sourceSha256, string? outputSha256, string recipeSha256, int approvedCount, int rejectedCount, int verifiedAbsentCount, int residualCount, int inconclusiveCount, bool verified, IReadOnlyList<int> pages, IReadOnlyList<PdfRedactionEvidenceIssue> issues, string encryptionPolicy, bool ocrUsed, IReadOnlyList<string> ocrProviders, bool ocrPostVerificationPerformed = false, int ocrResidualCandidateCount = 0, int sourceSignatureCount = 0, int outputSignatureCount = 0, string signaturePolicy = "RejectSignedSource", string? outputSigner = null, bool signatureCryptographicallyVerified = false, IReadOnlyList<string>? externalValidators = null) {
        SourceSha256 = sourceSha256; OutputSha256 = outputSha256; RecipeSha256 = recipeSha256; ApprovedCount = approvedCount; RejectedCount = rejectedCount;
        VerifiedAbsentCount = verifiedAbsentCount; ResidualCount = residualCount; InconclusiveCount = inconclusiveCount; Verified = verified; AffectedPageNumbers = pages;
        Issues = issues; EncryptionPolicy = encryptionPolicy; OcrUsed = ocrUsed; OcrProviders = ocrProviders; OcrPostVerificationPerformed = ocrPostVerificationPerformed; OcrResidualCandidateCount = ocrResidualCandidateCount;
        SourceSignatureCount = sourceSignatureCount; OutputSignatureCount = outputSignatureCount; SignaturePolicy = signaturePolicy; OutputSigner = outputSigner; SignatureCryptographicallyVerified = signatureCryptographicallyVerified; ExternalValidators = externalValidators ?? Array.Empty<string>();
    }
    /// <summary>Exact source fingerprint.</summary>
    public string SourceSha256 { get; }
    /// <summary>Output fingerprint, when an output was inspected.</summary>
    public string? OutputSha256 { get; }
    /// <summary>Canonical recipe fingerprint.</summary>
    public string RecipeSha256 { get; }
    /// <summary>Approved candidate count.</summary>
    public int ApprovedCount { get; }
    /// <summary>Rejected candidate count.</summary>
    public int RejectedCount { get; }
    /// <summary>Reviewed content items proven absent.</summary>
    public int VerifiedAbsentCount { get; }
    /// <summary>Residual reviewed content count.</summary>
    public int ResidualCount { get; }
    /// <summary>Inconclusive reviewed content count.</summary>
    public int InconclusiveCount { get; }
    /// <summary>Whether all configured proof checks passed.</summary>
    public bool Verified { get; }
    /// <summary>Affected pages.</summary>
    public IReadOnlyList<int> AffectedPageNumbers { get; }
    /// <summary>Stable issue codes without markers or extracted text.</summary>
    public IReadOnlyList<PdfRedactionEvidenceIssue> Issues { get; }
    /// <summary>Applied encrypted-document policy, or NoRewrite when the source was emitted byte-for-byte.</summary>
    public string EncryptionPolicy { get; }
    /// <summary>Whether OCR participated in candidate discovery.</summary>
    public bool OcrUsed { get; }
    /// <summary>Distinct provider identifiers reported by OCR.</summary>
    public IReadOnlyList<string> OcrProviders { get; }
    /// <summary>Whether the configured OCR provider also inspected the rewritten artifact.</summary>
    public bool OcrPostVerificationPerformed { get; }
    /// <summary>OCR candidates that still intersect approved areas after rewriting.</summary>
    public int OcrResidualCandidateCount { get; }
    /// <summary>Signature definitions observed on the exact source.</summary>
    public int SourceSignatureCount { get; }
    /// <summary>Signature definitions validated on the final derivative.</summary>
    public int OutputSignatureCount { get; }
    /// <summary>Applied signed-source and derivative policy.</summary>
    public string SignaturePolicy { get; }
    /// <summary>Stable signer implementation name, when the workflow signed the derivative.</summary>
    public string? OutputSigner { get; }
    /// <summary>Whether a caller-provided cryptographic validator verified signature math and digest.</summary>
    public bool SignatureCryptographicallyVerified { get; }
    /// <summary>Independent validator identifiers that inspected the final artifact.</summary>
    public IReadOnlyList<string> ExternalValidators { get; }
}

/// <summary>Operational plan, application, or verification result. Persist <see cref="PdfRedactionWorkflowRecord"/> when host paths and request identifiers must be omitted.</summary>
public sealed class PdfRedactionWorkflowResult {
    /// <summary>Stable result schema.</summary>
    public const string CurrentSchema = "officeimo.pdf.redaction.result.v1";
    /// <summary>Stable plan-only schema.</summary>
    public const string PlanSchema = "officeimo.pdf.redaction.plan.v1";
    internal PdfRedactionWorkflowResult(string requestId, PdfRedactionWorkflowMode mode, OfficeWorkflowStatus status, string summary, string sourceSha256, string recipeSha256, IReadOnlyList<PdfRedactionWorkflowCandidate> candidates, string? outputPath, string? evidencePath, PdfRedactionWorkflowEvidence? evidence, IReadOnlyList<OfficeWorkflowDiagnostic> diagnostics) {
        RequestId = requestId; Mode = mode; Status = status; Summary = summary; SourceSha256 = sourceSha256; RecipeSha256 = recipeSha256;
        Candidates = candidates; OutputPath = outputPath; EvidencePath = evidencePath; Evidence = evidence; Diagnostics = diagnostics;
    }
    /// <summary>Schema identifier.</summary>
    public string Schema => Mode == PdfRedactionWorkflowMode.PlanOnly ? PlanSchema : CurrentSchema;
    /// <summary>Caller request identifier.</summary>
    public string RequestId { get; }
    /// <summary>Executed mode.</summary>
    public PdfRedactionWorkflowMode Mode { get; }
    /// <summary>Terminal status.</summary>
    public OfficeWorkflowStatus Status { get; }
    /// <summary>Human-readable outcome without sensitive matched text.</summary>
    public string Summary { get; }
    /// <summary>Exact source fingerprint.</summary>
    public string SourceSha256 { get; }
    /// <summary>Canonical recipe fingerprint.</summary>
    public string RecipeSha256 { get; }
    /// <summary>Privacy-safe review candidates.</summary>
    public IReadOnlyList<PdfRedactionWorkflowCandidate> Candidates { get; }
    /// <summary>Published output path.</summary>
    public string? OutputPath { get; }
    /// <summary>Published JSON evidence path.</summary>
    public string? EvidencePath { get; }
    /// <summary>Post-rewrite evidence.</summary>
    public PdfRedactionWorkflowEvidence? Evidence { get; }
    /// <summary>Structured diagnostics.</summary>
    public IReadOnlyList<OfficeWorkflowDiagnostic> Diagnostics { get; }
    /// <summary>True when completed.</summary>
    public bool Succeeded => Status == OfficeWorkflowStatus.Completed;
}

/// <summary>Privacy-safe persisted plan or evidence record without host paths or caller correlation identifiers.</summary>
public sealed class PdfRedactionWorkflowRecord {
    internal PdfRedactionWorkflowRecord(PdfRedactionWorkflowResult result) {
        Schema = result.Schema;
        Mode = result.Mode;
        Status = result.Status;
        Summary = result.Summary;
        SourceSha256 = result.SourceSha256;
        RecipeSha256 = result.RecipeSha256;
        Candidates = result.Candidates;
        Evidence = result.Evidence;
        Diagnostics = result.Diagnostics;
    }

    /// <summary>Stable plan or result schema.</summary>
    public string Schema { get; }
    /// <summary>Executed mode.</summary>
    public PdfRedactionWorkflowMode Mode { get; }
    /// <summary>Terminal status.</summary>
    public OfficeWorkflowStatus Status { get; }
    /// <summary>Human-readable privacy-safe outcome.</summary>
    public string Summary { get; }
    /// <summary>Exact source fingerprint.</summary>
    public string SourceSha256 { get; }
    /// <summary>Canonical recipe fingerprint.</summary>
    public string RecipeSha256 { get; }
    /// <summary>Privacy-safe review candidates.</summary>
    public IReadOnlyList<PdfRedactionWorkflowCandidate> Candidates { get; }
    /// <summary>Post-rewrite evidence.</summary>
    public PdfRedactionWorkflowEvidence? Evidence { get; }
    /// <summary>Structured privacy-safe diagnostics.</summary>
    public IReadOnlyList<OfficeWorkflowDiagnostic> Diagnostics { get; }
}

/// <summary>Result of one bounded atomic redaction batch.</summary>
public sealed class PdfRedactionBatchResult {
    /// <summary>Stable batch schema.</summary>
    public const string CurrentSchema = "officeimo.pdf.redaction.batch.v1";
    internal PdfRedactionBatchResult(OfficeWorkflowStatus status, IReadOnlyList<PdfRedactionWorkflowResult> items, bool publishedAtomically, string summary) {
        Status = status; Items = items; PublishedAtomically = publishedAtomically; Summary = summary;
    }
    /// <summary>Schema identifier.</summary>
    public string Schema => CurrentSchema;
    /// <summary>Batch status.</summary>
    public OfficeWorkflowStatus Status { get; }
    /// <summary>Per-item results in request order.</summary>
    public IReadOnlyList<PdfRedactionWorkflowResult> Items { get; }
    /// <summary>True when all requested output artifacts were published as one successful batch transaction.</summary>
    public bool PublishedAtomically { get; }
    /// <summary>Human-readable batch outcome.</summary>
    public string Summary { get; }
}
