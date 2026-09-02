namespace OfficeIMO.Workflows;

/// <summary>Before or after snapshot used by PDF document-health reports.</summary>
public sealed class PdfHealthSnapshot {
    internal PdfHealthSnapshot(
        long sizeBytes,
        int pageCount,
        string? version,
        bool canRead,
        bool canRewrite,
        bool hasEncryption,
        bool hasSignatures,
        bool hasTaggedContent,
        bool hasActiveContent,
        bool hasEmbeddedFiles,
        int repairCount,
        int detectionOnlyCount,
        IReadOnlyList<string> diagnostics) {
        SizeBytes = sizeBytes;
        PageCount = pageCount;
        Version = version;
        CanRead = canRead;
        CanRewrite = canRewrite;
        HasEncryption = hasEncryption;
        HasSignatures = hasSignatures;
        HasTaggedContent = hasTaggedContent;
        HasActiveContent = hasActiveContent;
        HasEmbeddedFiles = hasEmbeddedFiles;
        RepairCount = repairCount;
        DetectionOnlyCount = detectionOnlyCount;
        Diagnostics = diagnostics.ToArray();
    }

    /// <summary>Artifact size.</summary>
    public long SizeBytes { get; }
    /// <summary>Readable page count.</summary>
    public int PageCount { get; }
    /// <summary>Effective PDF version when available.</summary>
    public string? Version { get; }
    /// <summary>Whether OfficeIMO.Pdf can read the document.</summary>
    public bool CanRead { get; }
    /// <summary>Whether a general full rewrite is available.</summary>
    public bool CanRewrite { get; }
    /// <summary>Whether encryption is present.</summary>
    public bool HasEncryption { get; }
    /// <summary>Whether signature markers are present.</summary>
    public bool HasSignatures { get; }
    /// <summary>Whether tagged content is present.</summary>
    public bool HasTaggedContent { get; }
    /// <summary>Whether active content is present.</summary>
    public bool HasActiveContent { get; }
    /// <summary>Whether embedded files are present.</summary>
    public bool HasEmbeddedFiles { get; }
    /// <summary>Recovered parser defect count.</summary>
    public int RepairCount { get; }
    /// <summary>Detected-only parser defect count.</summary>
    public int DetectionOnlyCount { get; }
    /// <summary>Read, rewrite, or repair diagnostics.</summary>
    public IReadOnlyList<string> Diagnostics { get; }
}

/// <summary>Explicit before/after evidence for a PDF health operation.</summary>
public sealed class PdfHealthReport {
    internal PdfHealthReport(
        OfficeWorkflowOperation operation,
        PdfHealthSnapshot before,
        PdfHealthSnapshot? after,
        string summary,
        bool verified,
        IReadOnlyDictionary<string, string>? metrics = null) {
        Operation = operation;
        Before = before;
        After = after;
        Summary = summary;
        Verified = verified;
        Metrics = new Dictionary<string, string>(metrics ?? new Dictionary<string, string>(), StringComparer.Ordinal);
    }

    /// <summary>Health operation.</summary>
    public OfficeWorkflowOperation Operation { get; }
    /// <summary>Source snapshot.</summary>
    public PdfHealthSnapshot Before { get; }
    /// <summary>Generated artifact snapshot, when the operation produced a PDF.</summary>
    public PdfHealthSnapshot? After { get; }
    /// <summary>User-facing outcome.</summary>
    public string Summary { get; }
    /// <summary>Whether operation-specific postconditions were verified.</summary>
    public bool Verified { get; }
    /// <summary>Operation-specific evidence.</summary>
    public IReadOnlyDictionary<string, string> Metrics { get; }
}

/// <summary>Completed, cancelled, or failed workflow result.</summary>
public sealed class OfficeWorkflowResult {
    internal OfficeWorkflowResult(
        string requestId,
        OfficeWorkflowOperation operation,
        OfficeWorkflowStatus status,
        OfficeWorkflowFailureKind failureKind,
        string? outputPath,
        long inputBytes,
        long outputBytes,
        TimeSpan duration,
        string summary,
        IReadOnlyList<OfficeWorkflowDiagnostic> diagnostics,
        PdfHealthReport? healthReport = null) {
        RequestId = requestId;
        Operation = operation;
        Status = status;
        FailureKind = failureKind;
        OutputPath = outputPath;
        InputBytes = inputBytes;
        OutputBytes = outputBytes;
        Duration = duration;
        Summary = summary;
        Diagnostics = diagnostics.ToArray();
        HealthReport = healthReport;
    }

    /// <summary>Caller-provided request identifier.</summary>
    public string RequestId { get; }
    /// <summary>Executed operation.</summary>
    public OfficeWorkflowOperation Operation { get; }
    /// <summary>Terminal state.</summary>
    public OfficeWorkflowStatus Status { get; }
    /// <summary>Stable failure category, or <see cref="OfficeWorkflowFailureKind.None"/> when no failure occurred.</summary>
    public OfficeWorkflowFailureKind FailureKind { get; }
    /// <summary>Final published output path, including a renamed destination when applicable.</summary>
    public string? OutputPath { get; }
    /// <summary>Primary input size.</summary>
    public long InputBytes { get; }
    /// <summary>Published artifact size, or zero for report-only operations.</summary>
    public long OutputBytes { get; }
    /// <summary>Total execution duration.</summary>
    public TimeSpan Duration { get; }
    /// <summary>User-facing outcome.</summary>
    public string Summary { get; }
    /// <summary>Structured fidelity, proof, warning, and failure diagnostics.</summary>
    public IReadOnlyList<OfficeWorkflowDiagnostic> Diagnostics { get; }
    /// <summary>Typed PDF health evidence when applicable.</summary>
    public PdfHealthReport? HealthReport { get; }
    /// <summary>True only for successfully completed requests.</summary>
    public bool Succeeded => Status == OfficeWorkflowStatus.Completed;
}
