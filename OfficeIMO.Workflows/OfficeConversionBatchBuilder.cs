namespace OfficeIMO.Workflows;

/// <summary>Fluent batch configuration using the same routes, profiles, renderer settings and publication policies as a single conversion.</summary>
public sealed class OfficeConversionBatchBuilder {
    private readonly OfficeConversionBatchRequest _request;
    internal OfficeConversionBatchBuilder(OfficeConversionBatchRequest request) => _request = request;
    /// <summary>Sets the destination tree and target extension. The default target is PDF.</summary>
    public OfficeConversionBatchBuilder ToDirectory(string outputDirectory, string targetExtension = ".pdf") {
        _request.OutputDirectory = outputDirectory; _request.TargetExtension = targetExtension; return this;
    }
    /// <summary>Selects an explicit executable route when source extensions are ambiguous.</summary>
    public OfficeConversionBatchBuilder Via(string routeId) { _request.ConversionRouteId = routeId; return this; }
    /// <summary>Captures existing format-specific renderer options.</summary>
    public OfficeConversionBatchBuilder WithConversionOptions(OfficeWorkflowConversionOptions options) {
        _request.ConversionOptions = options?.Clone() ?? throw new ArgumentNullException(nameof(options)); return this;
    }
    /// <summary>Enables hash-verified completion and restartable publication. Checkpoint jobs require the Fail conflict policy.</summary>
    public OfficeConversionBatchBuilder WithCheckpoint(string checkpointDirectory, bool retryFailed = false) {
        _request.CheckpointDirectory = checkpointDirectory; _request.RetryFailed = retryFailed; return this;
    }
    /// <summary>Sets the cross-format output intent.</summary>
    public OfficeConversionBatchBuilder WithProfile(OfficeWorkflowOutputProfile profile) { _request.OutputProfile = profile; return this; }
    /// <summary>Sets the runtime PDF input password without storing it in checkpoints.</summary>
    public OfficeConversionBatchBuilder WithPdfPassword(string? password) { _request.PdfPassword = password; return this; }
    /// <summary>Sets the bounded parallel execution budget.</summary>
    public OfficeConversionBatchBuilder WithConcurrency(int maximumConcurrency) { _request.MaximumConcurrency = maximumConcurrency; return this; }
    /// <summary>Sets per-document input and output byte limits and the maximum discovered file count.</summary>
    public OfficeConversionBatchBuilder WithLimits(long maximumInputBytes, long maximumOutputBytes, int maximumFiles = 1_000_000) {
        _request.MaximumInputBytes = maximumInputBytes; _request.MaximumOutputBytes = maximumOutputBytes;
        _request.MaximumFiles = maximumFiles; return this;
    }
    /// <summary>Sets ordinary batch destination behavior; durable jobs require Fail.</summary>
    public OfficeConversionBatchBuilder OnConflict(OfficeWorkflowConflictPolicy policy) { _request.ConflictPolicy = policy; return this; }
    /// <summary>Restricts directory selection to source extensions and optional descendants.</summary>
    public OfficeConversionBatchBuilder SelectExtensions(bool recursive, params string[] extensions) {
        _request.Recursive = recursive; _request.SourceExtensions = extensions?.ToArray() ?? throw new ArgumentNullException(nameof(extensions)); return this;
    }
    /// <summary>Creates an independent request snapshot.</summary>
    public OfficeConversionBatchRequest Build() => _request with {
        InputPaths = _request.InputPaths?.ToArray(), SourceExtensions = _request.SourceExtensions?.ToArray(),
        ConversionOptions = _request.ConversionOptions.Clone()
    };
    /// <summary>Executes this batch. Item callbacks may arrive concurrently; the summary remains bounded.</summary>
    public Task<OfficeConversionBatchResult> RunAsync(IOfficeWorkflowRunner? runner = null,
        IProgress<OfficeConversionBatchItemResult>? progress = null, CancellationToken cancellationToken = default,
        IOfficeWorkflowPublicationGuard? publicationGuard = null) =>
        (runner ?? new OfficeWorkflowRunner()).RunBatchAsync(Build(), progress, cancellationToken, publicationGuard);
}
