using OfficeIMO.Provenance;

namespace OfficeIMO.Workflows;

/// <summary>Completed, cancelled, or failed provenance workflow result.</summary>
public sealed class OfficeProvenanceWorkflowResult {
    internal OfficeProvenanceWorkflowResult(
        string requestId,
        OfficeProvenanceWorkflowOperation operation,
        OfficeWorkflowStatus status,
        OfficeWorkflowFailureKind failureKind,
        string ownerPackage,
        string? outputPath,
        long inputBytes,
        long outputBytes,
        TimeSpan duration,
        string summary,
        IReadOnlyList<OfficeWorkflowDiagnostic> diagnostics,
        OfficeProvenanceReport? inspection = null,
        OfficeProvenanceAssessmentReport? assessment = null,
        OfficeProvenanceReport? before = null,
        OfficeProvenanceReport? after = null,
        IReadOnlyList<OfficeProvenanceChange>? changes = null,
        bool wasReserialized = false,
        bool wereInvalidatedSignaturesRemoved = false) {
        RequestId = requestId;
        Operation = operation;
        Status = status;
        FailureKind = failureKind;
        OwnerPackage = ownerPackage;
        OutputPath = outputPath;
        InputBytes = inputBytes;
        OutputBytes = outputBytes;
        Duration = duration;
        Summary = summary;
        Diagnostics = diagnostics.ToArray();
        Inspection = inspection;
        Assessment = assessment;
        Before = before;
        After = after;
        Changes = (changes ?? Array.Empty<OfficeProvenanceChange>()).ToArray();
        WasReserialized = wasReserialized;
        WereInvalidatedSignaturesRemoved = wereInvalidatedSignaturesRemoved;
    }

    /// <summary>Caller-provided request identifier.</summary>
    public string RequestId { get; }
    /// <summary>Executed operation.</summary>
    public OfficeProvenanceWorkflowOperation Operation { get; }
    /// <summary>Terminal execution state.</summary>
    public OfficeWorkflowStatus Status { get; }
    /// <summary>Stable failure category.</summary>
    public OfficeWorkflowFailureKind FailureKind { get; }
    /// <summary>Package that owned format-specific behavior.</summary>
    public string OwnerPackage { get; }
    /// <summary>Published removal artifact, when applicable.</summary>
    public string? OutputPath { get; }
    /// <summary>Input file size.</summary>
    public long InputBytes { get; }
    /// <summary>Published output size.</summary>
    public long OutputBytes { get; }
    /// <summary>Elapsed workflow duration.</summary>
    public TimeSpan Duration { get; }
    /// <summary>User-facing outcome.</summary>
    public string Summary { get; }
    /// <summary>Structured orchestration diagnostics.</summary>
    public IReadOnlyList<OfficeWorkflowDiagnostic> Diagnostics { get; }
    /// <summary>Structural report for an inspect operation.</summary>
    public OfficeProvenanceReport? Inspection { get; }
    /// <summary>Combined evidence for an assess operation.</summary>
    public OfficeProvenanceAssessmentReport? Assessment { get; }
    /// <summary>Structural evidence before removal.</summary>
    public OfficeProvenanceReport? Before { get; }
    /// <summary>Reopened structural evidence after removal.</summary>
    public OfficeProvenanceReport? After { get; }
    /// <summary>Format-native mutations applied in source order.</summary>
    public IReadOnlyList<OfficeProvenanceChange> Changes { get; }
    /// <summary>Whether the owner changed at least one carrier.</summary>
    public bool WasChanged => Changes.Count != 0;
    /// <summary>Whether the owner serialized a container rather than copying bytes around carriers.</summary>
    public bool WasReserialized { get; }
    /// <summary>Whether explicitly authorized invalidated signatures were removed.</summary>
    public bool WereInvalidatedSignaturesRemoved { get; }
    /// <summary>Whether the workflow completed successfully.</summary>
    public bool Succeeded => Status == OfficeWorkflowStatus.Completed;
}
