using OfficeIMO.Internal;

namespace OfficeIMO.Workflows;

public sealed partial class OfficeWorkflowRunner {
    private sealed record PublishedProviderFile(string Path, long SizeBytes);
    private sealed record ProviderFileBatch(OfficeWorkflowStatus Status, string? Failure,
        IReadOnlyList<PublishedProviderFile> Files, IReadOnlyList<OfficeWorkflowOutputRecovery> Recoveries);

    // Resolves the complete set without writes, then publishes and verifies each child with shared recovery.
    private static async Task<ProviderFileBatch> PublishProviderFilesAsync(IReadOnlyList<string> stagedPaths,
        OfficeWorkflowDirectoryOutput directory, long maximumOutputBytes, IOfficeWorkflowPublicationGuard? publicationGuard,
        List<OfficeWorkflowDiagnostic> diagnostics, CancellationToken token) {
        var files = new List<PublishedProviderFile>();
        var recoveries = new List<OfficeWorkflowOutputRecovery>();
        var destinations = new List<OfficeWorkflowDirectoryOutputFile>();
        try {
            foreach (string path in stagedPaths) {
                token.ThrowIfCancellationRequested();
                string name = Path.GetFileName(path);
                var destination = await directory.ResolveFile(name, token).ConfigureAwait(false)
                    ?? throw new IOException("The provider could not resolve an output file.");
                if (!string.Equals(destination.Output.Name, name, StringComparison.Ordinal))
                    throw new IOException("The provider changed the requested output filename.");
                if (destination.Output.PrepareDestination is null && destinations.Any(previous =>
                    previous.Output.PrepareDestination is null && OfficeStorageIdentity.AreEquivalent(previous.Location, destination.Location)))
                    throw new IOException("The provider mapped multiple output files to the same destination.");
                destinations.Add(destination);
            }
            for (int index = 0; index < stagedPaths.Count; index++) {
                token.ThrowIfCancellationRequested();
                string path = stagedPaths[index];
                var destination = destinations[index];
                var guard = new WorkflowScopedSourcePublicationGuard(
                    new DistinctWorkflowOutputPublicationGuard(publicationGuard, () => files.Select(file => file.Path)), [], [], destination.Output);
                var outcome = await PublishProviderArtifactAsync(path, destination.Location, destination.Output,
                    maximumOutputBytes, guard, () => File.Delete(path), diagnostics, token).ConfigureAwait(false);
                if (outcome.Recovery is not null) recoveries.Add(outcome.Recovery);
                if (outcome.Status != OfficeWorkflowStatus.Completed)
                    return new(outcome.Status, outcome.Summary, files, recoveries);
                files.Add(new(outcome.PublishedLocation, outcome.OutputBytes));
            }
            return new(OfficeWorkflowStatus.Completed, null, files, recoveries);
        } catch (Exception error) when (error is not OutOfMemoryException and not StackOverflowException) {
            var status = error is OperationCanceledException && token.IsCancellationRequested
                ? OfficeWorkflowStatus.Cancelled : OfficeWorkflowStatus.Failed;
            diagnostics.Add(new("ProviderFileBatchStopped", error.Message,
                status == OfficeWorkflowStatus.Cancelled ? OfficeWorkflowDiagnosticSeverity.Information : OfficeWorkflowDiagnosticSeverity.Error, "publish"));
            return new(status, error.Message, files, recoveries);
        }
    }
}
