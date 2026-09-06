using System.Diagnostics;
using OfficeIMO.Drawing;
using OfficeIMO.Internal;

namespace OfficeIMO.Workflows;

public sealed partial class OfficeWorkflowRunner {
    private static async Task<PdfPageImageExportResult> PublishProviderImagesAsync(ValidatedImageExportRequest request,
        OfficeImageExportBatchSaveResult saved, int[] pageNumbers, long inputBytes, Stopwatch stopwatch,
        List<OfficeWorkflowDiagnostic> diagnostics, CancellationToken token) {
        var files = new List<PdfPageImageFile>();
        var recoveries = new List<OfficeWorkflowOutputRecovery>();
        var destinations = new List<OfficeWorkflowDirectoryOutputFile>();
        long outputBytes = 0;
        OfficeWorkflowStatus status = OfficeWorkflowStatus.Completed;
        string? failure = null;
        try {
            // Resolve the whole batch before writing. A resolver is a read-only provider operation.
            foreach (var file in saved.Files) {
                token.ThrowIfCancellationRequested();
                string name = Path.GetFileName(file.Path);
                var destination = await request.DirectoryOutput!.ResolveFile(name, token).ConfigureAwait(false)
                    ?? throw new IOException("The provider could not resolve an output file.");
                if (!string.Equals(destination.Output.Name, name, StringComparison.Ordinal))
                    throw new IOException("The provider changed the requested output filename.");
                if (destination.Output.PrepareDestination is null && destinations.Any(previous =>
                    previous.Output.PrepareDestination is null && OfficeStorageIdentity.AreEquivalent(previous.Location, destination.Location)))
                    throw new IOException("The provider mapped multiple output files to the same destination.");
                destinations.Add(destination);
            }
            for (int index = 0; index < saved.Files.Count; index++) {
                token.ThrowIfCancellationRequested();
                var file = saved.Files[index];
                var destination = destinations[index];
                // Keep a local destination's provider scope active while the captured source and host guards run.
                var guard = new WorkflowScopedSourcePublicationGuard(
                    new DistinctImagePublicationGuard(request.PublicationGuard, files), [], [], destination.Output);
                var outcome = await PublishProviderArtifactAsync(file.Path, destination.Location, destination.Output,
                    request.Limits.MaximumOutputBytes, guard, () => File.Delete(file.Path), diagnostics, token).ConfigureAwait(false);
                if (outcome.Recovery is not null) recoveries.Add(outcome.Recovery);
                if (outcome.Status != OfficeWorkflowStatus.Completed) {
                    status = outcome.Status;
                    failure = outcome.Summary;
                    break;
                }
                outputBytes = checked(outputBytes + outcome.OutputBytes);
                files.Add(new(pageNumbers[index], outcome.PublishedLocation, file.Format, file.Width, file.Height, outcome.OutputBytes));
            }
        } catch (Exception error) when (error is not OutOfMemoryException and not StackOverflowException) {
            status = error is OperationCanceledException && token.IsCancellationRequested
                ? OfficeWorkflowStatus.Cancelled : OfficeWorkflowStatus.Failed;
            failure = error.Message;
            diagnostics.Add(new("ProviderImageBatchStopped", failure,
                status == OfficeWorkflowStatus.Cancelled ? OfficeWorkflowDiagnosticSeverity.Information : OfficeWorkflowDiagnosticSeverity.Error, "publish"));
        }
        string summary = status == OfficeWorkflowStatus.Completed
            ? $"Exported and verified {files.Count:N0} page images in the provider folder."
            : $"Verified {files.Count:N0} of {saved.Files.Count:N0} page images. Previously verified files remain in the destination. {failure}";
        return new(request.Id, status,
            status is OfficeWorkflowStatus.Completed or OfficeWorkflowStatus.Cancelled ? OfficeWorkflowFailureKind.None : OfficeWorkflowFailureKind.OutputFailed,
            files.Count > 0 ? request.OutputDirectory : null, inputBytes, outputBytes, stopwatch.Elapsed, summary, files, diagnostics, recoveries);
    }

    private sealed class DistinctImagePublicationGuard(IOfficeWorkflowPublicationGuard? host,
        IReadOnlyList<PdfPageImageFile> published) : IOfficeWorkflowPublicationGuard {
        public async ValueTask<bool> CanPublishAsync(string path, bool isDirectory, CancellationToken token) {
            if (published.Any(file => OfficeStorageIdentity.AreEquivalent(file.Path, path))) return false;
            if (host is not null && !await host.CanPublishAsync(path, isDirectory, token).ConfigureAwait(false)) return false;
            token.ThrowIfCancellationRequested();
            return !published.Any(file => OfficeStorageIdentity.AreEquivalent(file.Path, path));
        }
    }
}
