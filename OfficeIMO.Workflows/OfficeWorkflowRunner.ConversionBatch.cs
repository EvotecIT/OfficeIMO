namespace OfficeIMO.Workflows;

public sealed partial class OfficeWorkflowRunner {
    /// <summary>Converts an incrementally discovered directory or explicit files through the existing routes.
    /// Optional checkpoints retain verified completion and publication intent. Item callbacks may run concurrently.</summary>
    public Task<OfficeConversionBatchResult> RunBatchAsync(OfficeConversionBatchRequest request,
        IProgress<OfficeConversionBatchItemResult>? progress = null, CancellationToken cancellationToken = default,
        IOfficeWorkflowPublicationGuard? publicationGuard = null) =>
        OfficeConversionBatchExecutor.RunAsync(this, request, progress, cancellationToken, publicationGuard);
}
