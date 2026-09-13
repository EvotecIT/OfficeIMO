using OfficeIMO.Provenance;

namespace OfficeIMO.Workflows;

/// <summary>Operations exposed by the cross-format provenance workflow.</summary>
public enum OfficeProvenanceWorkflowOperation {
    /// <summary>Inspect standards-defined structural provenance carriers.</summary>
    Inspect,
    /// <summary>Combine structural, cryptographic, text-integrity, and provider-specific evidence.</summary>
    Assess,
    /// <summary>Remove selected standards-defined provenance carriers through the owning format API.</summary>
    Remove
}
/// <summary>One typed cross-format provenance request.</summary>
public sealed class OfficeProvenanceWorkflowRequest {
    /// <summary>Caller-provided identifier used by progress and batch results.</summary>
    public string Id { get; set; } = Guid.NewGuid().ToString("N");

    /// <summary>Requested provenance operation.</summary>
    public OfficeProvenanceWorkflowOperation Operation { get; set; }

    /// <summary>Input asset path.</summary>
    public required string InputPath { get; set; }

    /// <summary>Output asset path for removal. When omitted, a sibling provenance-cleaned name is used.</summary>
    public string? OutputPath { get; set; }

    /// <summary>Conflict behavior used when publishing a removal artifact.</summary>
    public OfficeWorkflowConflictPolicy ConflictPolicy { get; set; } = OfficeWorkflowConflictPolicy.Rename;

    /// <summary>Structural inspection limits used by <see cref="OfficeProvenanceWorkflowOperation.Inspect"/>.</summary>
    public OfficeProvenanceOptions Inspection { get; set; } = new();

    /// <summary>Combined assessment options used by <see cref="OfficeProvenanceWorkflowOperation.Assess"/>.</summary>
    public OfficeProvenanceAssessmentOptions Assessment { get; set; } = new();

    /// <summary>Selective removal policy used by <see cref="OfficeProvenanceWorkflowOperation.Remove"/>.</summary>
    public OfficeProvenanceRemovalOptions Removal { get; set; } = new();

    /// <summary>Shared input and output byte limits.</summary>
    public OfficeWorkflowLimits Limits { get; set; } = new();

    internal SortedSet<string>? BatchBlockedOutputIdentities { get; set; }
    internal string? BatchOwnReservedOutputIdentity { get; set; }
}

/// <summary>Bounds sequential provenance batch execution.</summary>
public sealed class OfficeProvenanceWorkflowBatchOptions {
    /// <summary>Maximum number of materialized requests. Defaults to 256.</summary>
    public int MaximumRequests { get; set; } = 256;

    /// <summary>Whether execution continues after a failed request. Defaults to true.</summary>
    public bool ContinueOnFailure { get; set; } = true;

    internal OfficeProvenanceWorkflowBatchOptions CloneAndValidate() {
        if (MaximumRequests <= 0 || MaximumRequests > 10_000) {
            throw new ArgumentOutOfRangeException(nameof(MaximumRequests), "MaximumRequests must be between 1 and 10,000.");
        }
        return new OfficeProvenanceWorkflowBatchOptions {
            MaximumRequests = MaximumRequests,
            ContinueOnFailure = ContinueOnFailure
        };
    }
}

/// <summary>Runs reusable cross-format provenance workflows.</summary>
public interface IOfficeProvenanceWorkflowRunner {
    /// <summary>Runs one provenance request.</summary>
    Task<OfficeProvenanceWorkflowResult> RunProvenanceAsync(
        OfficeProvenanceWorkflowRequest request,
        IProgress<OfficeWorkflowProgress>? progress = null,
        CancellationToken cancellationToken = default);

    /// <summary>Runs a bounded provenance batch sequentially.</summary>
    Task<IReadOnlyList<OfficeProvenanceWorkflowResult>> RunProvenanceBatchAsync(
        IEnumerable<OfficeProvenanceWorkflowRequest> requests,
        OfficeProvenanceWorkflowBatchOptions? options = null,
        IProgress<OfficeWorkflowProgress>? progress = null,
        CancellationToken cancellationToken = default);
}
