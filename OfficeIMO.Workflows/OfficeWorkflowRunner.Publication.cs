namespace OfficeIMO.Workflows;

public sealed partial class OfficeWorkflowRunner {
        private static Task<string> PublishAsync(string stagingPath, string requestedPath,
        OfficePublicationPolicy policy, IOfficePublicationGuard? guard, CancellationToken cancellationToken) =>
        OfficeFilePublisher.PublishAsync(stagingPath, requestedPath, policy, guard, cancellationToken);
    private static async Task<bool> CanPublishAsync(
        IOfficePublicationGuard? guard, string path, bool isDirectory, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        bool allowed = guard is null || await guard.CanPublishAsync(path, isDirectory, cancellationToken).ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
        return allowed;
    }

    private static async ValueTask EnsurePublicationAllowedAsync(
        IOfficePublicationGuard? guard, string path, bool isDirectory, CancellationToken cancellationToken) {
        if (!await CanPublishAsync(guard, path, isDirectory, cancellationToken).ConfigureAwait(false)) {
            throw new IOException("The output destination is in use by the application: " + path);
        }
    }
}
