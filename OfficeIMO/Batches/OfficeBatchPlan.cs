namespace OfficeIMO;

/// <summary>One preflight observation. Execution revalidates current input and output policy.</summary>
public sealed record OfficeBatchPlanItem(string InputPath, string? OutputPath, string? RouteId, string Disposition, string Message);

/// <summary>Bounded preflight without rendering, destination reservations or checkpoint reuse.</summary>
public sealed class OfficeBatchPlan {
    /// <summary>Whether preflight found eligible input without blocking findings.</summary>
    public bool CanRun => Selected > 0 && Failed == 0;
    /// <summary>Discovered inputs.</summary>
    public long Discovered { get; internal set; }
    /// <summary>Inputs matching an executable route.</summary>
    public long Selected { get; internal set; }
    /// <summary>Inputs excluded by the selection policy.</summary>
    public long Skipped { get; internal set; }
    /// <summary>Blocking input, configuration or discovery findings.</summary>
    public long Failed { get; internal set; }
    /// <summary>Proposed destinations already occupied or shared by selected input.</summary>
    public long Collisions { get; internal set; }
    /// <summary>At most 250 representative observations. Progress receives every input.</summary>
    public IReadOnlyList<OfficeBatchPlanItem> Items { get; internal set; } = [];
    /// <summary>Whether some observations were omitted to bound the plan.</summary>
    public bool ItemsTruncated { get; internal set; }
    /// <summary>At most 32 general preflight notes.</summary>
    public IReadOnlyList<string> Diagnostics { get; internal set; } = [];
}
