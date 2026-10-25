namespace OfficeIMO.Workflows;

/// <summary>One preflight observation. Output names are proposals; execution revalidates current files and collisions.</summary>
public sealed record OfficeConversionJobPlanItem(string InputPath, string? OutputPath, string? RouteId, string Disposition, string Message);

/// <summary>A bounded preflight summary. Planning does not render, reserve destinations or verify checkpoint artifacts.</summary>
public sealed class OfficeConversionJobPlan {
    /// <summary>Whether preflight found eligible inputs without blocking issues.</summary>
    public bool CanRun => Selected > 0 && Failed == 0;
    /// <summary>Discovered input count.</summary>
    public long Discovered { get; internal set; }
    /// <summary>Inputs matching executable routes.</summary>
    public long Selected { get; internal set; }
    /// <summary>Inputs excluded by extension or route selection.</summary>
    public long Skipped { get; internal set; }
    /// <summary>Blocking input, configuration or discovery issues.</summary>
    public long Failed { get; internal set; }
    /// <summary>Proposed paths already occupied or shared by selected inputs.</summary>
    public long Collisions { get; internal set; }
    /// <summary>At most 250 representative input observations. Progress callbacks receive each discovered item.</summary>
    public IReadOnlyList<OfficeConversionJobPlanItem> Items { get; internal set; } = [];
    /// <summary>Whether the observations were truncated.</summary>
    public bool ItemsTruncated { get; internal set; }
    /// <summary>At most 32 general preflight notes.</summary>
    public IReadOnlyList<string> Diagnostics { get; internal set; } = [];
}

/// <summary>Plans and runs saved jobs through the same configured workflow runner as interactive hosts.</summary>
public static class OfficeConversionJobRunner {
    /// <summary>Discovers inputs without output/checkpoint writes or secret resolution. Execute revalidates all source and output policy.</summary>
    public static Task<OfficeConversionJobPlan> PlanAsync(OfficeConversionJob job, string baseDirectory, IOfficeWorkflowRunner runner,
        IProgress<OfficeConversionJobPlanItem>? progress = null, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(runner);
        return OfficeConversionBatchExecutor.PlanAsync(runner, job.CreateRequest(baseDirectory), progress, cancellationToken);
    }

    /// <summary>Resolves admitted runtime secrets and delegates to the shared batch executor. Checkpoint and publication restrictions are unchanged.</summary>
    public static Task<OfficeConversionBatchResult> RunAsync(OfficeConversionJob job, string baseDirectory, IOfficeWorkflowRunner runner,
        OfficeConversionJobRuntimeOptions? runtime = null, IProgress<OfficeConversionBatchItemResult>? progress = null,
        CancellationToken cancellationToken = default, IOfficeWorkflowPublicationGuard? publicationGuard = null) {
        ArgumentNullException.ThrowIfNull(runner);
        return runner.RunBatchAsync(job.CreateRequest(baseDirectory, runtime ?? new OfficeConversionJobRuntimeOptions()), progress, cancellationToken, publicationGuard);
    }
}
