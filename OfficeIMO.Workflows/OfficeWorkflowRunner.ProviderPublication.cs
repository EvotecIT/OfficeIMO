using OfficeIMO.Core.Internal;

namespace OfficeIMO.Workflows;

public sealed partial class OfficeWorkflowRunner {
    private static async Task<ProviderPublicationOutcome> PublishProviderArtifactAsync(string stagingPath, string destination,
        OfficeWorkflowStreamOutput provider, long maximumBytes, IOfficeWorkflowPublicationGuard? guard,
        Action cleanupStaging, List<OfficeWorkflowDiagnostic> diagnostics, CancellationToken token) {
        OfficeWorkflowOutputRecoveryStore.RecoveryLease? lease = null;
        OfficeWorkflowOutputRecovery? recovery = null;
        OfficeWorkflowStatus status = OfficeWorkflowStatus.Failed;
        long outputBytes = 0;
        string summary;
        try {
            byte[] bytes;
            await using (var source = new FileStream(stagingPath, FileMode.Open, FileAccess.Read, FileShare.Read)) {
                bytes = await OfficeStreamReader.ReadAllBytesAsync(source, token, maximumBytes).ConfigureAwait(false);
            }
            lease = await provider.RecoveryStore.CreateAsync(bytes, destination, provider.Name, token).ConfigureAwait(false);
            cleanupStaging();
            await EnsurePublicationAllowedAsync(guard, destination, false, token).ConfigureAwait(false);
            await OfficeStreamPublication.WriteVerifiedAsync(provider.OpenRead, provider.OpenWrite, bytes,
                expectedFingerprint: null, maximumBytes, token).ConfigureAwait(false);
            status = OfficeWorkflowStatus.Completed;
            outputBytes = bytes.LongLength;
            summary = "The provider output was saved and its contents were verified.";
            diagnostics.Add(new OfficeWorkflowDiagnostic("ProviderPublicationVerified",
                "The provider retained the complete output after the write stream closed. Atomic replacement and rollback are not available.",
                stage: "publish"));
        } catch (Exception error) when (error is not OutOfMemoryException and not StackOverflowException) {
            if (error.Data["OfficeIMO.Storage.OutputRecoveryCleanupFailed"] is string directory) {
                diagnostics.Add(new OfficeWorkflowDiagnostic("OutputRecoveryCleanupFailed", "Output recovery staging could not be removed: " + directory,
                    OfficeWorkflowDiagnosticSeverity.Warning, "cleanup"));
            }
            bool uncertain = OfficeStreamPublication.MayHaveChangedDestination(error);
            status = uncertain ? OfficeWorkflowStatus.Unconfirmed
                : error is OperationCanceledException && token.IsCancellationRequested ? OfficeWorkflowStatus.Cancelled : OfficeWorkflowStatus.Failed;
            if (uncertain) recovery = lease?.Recovery;
            summary = uncertain
                ? "The provider output could not be verified. Check the destination and open the recovery copy before trying again."
                : status == OfficeWorkflowStatus.Cancelled ? "Cancelled before the provider write began." : "Provider publication failed: " + error.Message;
            diagnostics.Add(new OfficeWorkflowDiagnostic(uncertain ? "ProviderPublicationUnconfirmed" : "ProviderPublicationFailed",
                summary, status == OfficeWorkflowStatus.Cancelled ? OfficeWorkflowDiagnosticSeverity.Information : OfficeWorkflowDiagnosticSeverity.Error,
                "publish"));
        } finally {
            lease?.Dispose();
        }
        if (lease is not null && recovery is null) {
            try { provider.RecoveryStore.Discard(lease.Recovery); }
            catch (Exception error) when (error is IOException or UnauthorizedAccessException) {
                recovery = lease.Recovery;
                diagnostics.Add(new OfficeWorkflowDiagnostic("OutputRecoveryRetained",
                    "The local recovery copy could not be removed: " + error.Message,
                    OfficeWorkflowDiagnosticSeverity.Warning, "cleanup"));
            }
        }
        return new(status, outputBytes, summary, recovery);
    }

    private sealed record ProviderPublicationOutcome(OfficeWorkflowStatus Status, long OutputBytes,
        string Summary, OfficeWorkflowOutputRecovery? Recovery);
}
