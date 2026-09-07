namespace OfficeIMO.Workflows;

public sealed partial class OfficeWorkflowRunner {
    /// <summary>Runs a batch sequentially so every request shares one predictable local resource budget.</summary>
    public async Task<IReadOnlyList<OfficeWorkflowResult>> RunBatchAsync(
        IEnumerable<OfficeWorkflowRequest> requests,
        IProgress<OfficeWorkflowProgress>? progress = null,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(requests);
        if (cancellationToken.IsCancellationRequested) return Array.Empty<OfficeWorkflowResult>();
        var batch = new List<PreparedRequest>();
        var selectedSources = new List<(string Location, OfficeWorkflowStreamInput? Stream)>();
        using (IEnumerator<OfficeWorkflowRequest> enumerator = requests.GetEnumerator()) {
            while (true) {
                if (cancellationToken.IsCancellationRequested) return Array.Empty<OfficeWorkflowResult>();
                if (!enumerator.MoveNext()) break;
                if (cancellationToken.IsCancellationRequested) return Array.Empty<OfficeWorkflowResult>();
                if (batch.Count >= MaximumBatchRequestCount) {
                    throw new InvalidOperationException(
                        $"A workflow batch cannot contain more than {MaximumBatchRequestCount:N0} requests.");
                }
                OfficeWorkflowRequest request = enumerator.Current
                    ?? throw new ArgumentException("Batch requests cannot contain null entries.", nameof(requests));
                batch.Add(PrepareRequest(request));
                AddSource(request.InputPath, request.InputStream);
                AddSource(request.ComparisonPath, request.ComparisonStream);
            }
        }

        var sources = selectedSources.GroupBy(item => item.Location, StringComparer.Ordinal)
            .Select(group => group.FirstOrDefault(item => item.Stream is not null) is { Stream: not null } scoped ? scoped : group.First()).ToArray();
        string[] protectedSources = sources.Select(item => item.Location!).Distinct(StringComparer.Ordinal).ToArray();
        WorkflowSourceAccess[] accesses = sources.Where(item => item.Stream is not null)
            .Select(item => new WorkflowSourceAccess(item.Location!, item.Stream!)).ToArray();
        var results = new List<OfficeWorkflowResult>(batch.Count);
        for (int i = 0; i < batch.Count; i++) {
            if (cancellationToken.IsCancellationRequested) break;
            PreparedRequest request = batch[i];
            if (request.Validated is { } validated) {
                try {
                    request = request with { Validated = validated with {
                        PublicationGuard = new WorkflowScopedSourcePublicationGuard(validated.PublicationGuard, protectedSources, accesses, validated.OutputStream,
                            allowMissingLocalSources: true)
                    } };
                } catch (Exception error) when (error is not OutOfMemoryException and not StackOverflowException) {
                    request = request with { ValidationException = error };
                }
            }
            int batchIndex = i;
            var batchProgress = progress is null
                ? null
                : new InlineProgress<OfficeWorkflowProgress>(item => progress.Report(new OfficeWorkflowProgress(
                    item.RequestId,
                    item.Stage,
                    $"{batchIndex + 1} of {batch.Count} · {item.Message}",
                    item.Fraction,
                    (batchIndex + item.Fraction) / Math.Max(1, batch.Count))));
            results.Add(await RunPreparedAsync(request, batchProgress, cancellationToken).ConfigureAwait(false));
        }
        return results;

        void AddSource(string? location, OfficeWorkflowStreamInput? stream) {
            if (string.IsNullOrWhiteSpace(location)) return;
            try { selectedSources.Add((OfficeIMO.Internal.OfficeStorageIdentity.Normalize(location), stream)); }
            catch (Exception error) when (error is ArgumentException or NotSupportedException or PathTooLongException) {
                // An invalid location cannot identify a destination; its own request retains the validation failure.
            }
        }
    }

}
