using OfficeIMO.IWork;

namespace OfficeIMO.Workflows.IWork;

/// <summary>Per-request read bounds and acceptance choices for an Apple workflow conversion.</summary>
public sealed class IWorkWorkflowSettings : IOfficeWorkflowConversionSettings {
    /// <summary>Source package and materialization limits. The workflow input limit also applies.</summary>
    public IWorkReadOptions ReadOptions { get; set; } = new() { PreserveSourceRecords = false };
    /// <summary>Representation, partial reconstruction, preview coverage, and worksheet naming choices.</summary>
    public IWorkConversionOptions ConversionOptions { get; set; } = new() { RequireCompleteVisualCoverage = true };

    /// <summary>Validates and copies all mutable options for one request.</summary>
    public IWorkWorkflowSettings Clone() => new() {
        ReadOptions = (ReadOptions ?? throw new ArgumentException("iWork read options cannot be null.")).Clone(),
        ConversionOptions = (ConversionOptions ?? throw new ArgumentException("iWork conversion options cannot be null.")).Clone()
    };

    /// <inheritdoc />
    public IOfficeWorkflowConversionSettings Snapshot() => Clone();
}
