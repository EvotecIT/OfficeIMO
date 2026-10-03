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
        var selectedSources = new List<(string Location, OfficeWorkflowStreamInput? Stream, OfficeWorkflowDirectoryPackageInput? Package)>();
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
                PreparedRequest prepared = PrepareRequest(request);
                batch.Add(prepared);
                OfficeWorkflowStreamInput? packageStream = prepared.Validated?.InputStream is { SnapshotKind: OfficeWorkflowSourceSnapshotKind.DirectoryPackage } captured ? captured : null;
                AddSource(request.InputPath, request.InputStream ?? packageStream, request.InputDirectoryPackage);
                AddSource(request.ComparisonPath, request.ComparisonStream, null);
            }
        }

        var packages = selectedSources.Where(item => item.Package is not null)
            .GroupBy(item => item.Location, StringComparer.Ordinal).Select(group => group.First()).ToArray();
        var sources = selectedSources.Where(item => !packages.Any(package => package.Location == item.Location)).GroupBy(item => item.Location, StringComparer.Ordinal)
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
                    // Package roots must be inspected by their permission-aware owner, never as raw local paths.
                    // Reject another selected package as a destination before creating output staging inside it.
                    foreach (var package in packages) {
                        if (validated.OutputPath is not null && !await new ProviderPackagePublicationGuard(
                            package.Package!.SourcePublicationGuard, null, package.Location)
                            .CanPublishAsync(validated.OutputPath, false, cancellationToken).ConfigureAwait(false))
                            throw new IOException("The batch output is not separate from a selected directory package.");
                    }
                    foreach (var access in accesses.Where(access => access.IsDirectoryPackage)) {
                        if (validated.OutputPath is null) continue;
                        await using var scope = await access.OpenReadAsync(cancellationToken).ConfigureAwait(false);
                        if (access.SourceGuard is { } owner && !await owner.CanPublishAsync(validated.OutputPath, false, cancellationToken).ConfigureAwait(false))
                            throw new IOException("The batch output is not separate from a selected directory package.");
                    }
                    IOfficeWorkflowPublicationGuard? guard = validated.PublicationGuard;
                    foreach (var package in packages) guard = new ProviderPackagePublicationGuard(
                        package.Package!.SourcePublicationGuard, guard, package.Location);
                    request = request with { Validated = validated with {
                        PublicationGuard = new WorkflowScopedSourcePublicationGuard(guard, protectedSources, accesses,
                            validated.OutputStream, allowMissingLocalSources: true)
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

        void AddSource(string? location, OfficeWorkflowStreamInput? stream, OfficeWorkflowDirectoryPackageInput? package) {
            if (string.IsNullOrWhiteSpace(location)) return;
            try { selectedSources.Add((OfficeIMO.Internal.OfficeStorageIdentity.Normalize(location), stream, package)); }
            catch (Exception error) when (error is ArgumentException or NotSupportedException or PathTooLongException) {
                // An invalid location cannot identify a destination; its own request retains the validation failure.
            }
        }
    }

}
