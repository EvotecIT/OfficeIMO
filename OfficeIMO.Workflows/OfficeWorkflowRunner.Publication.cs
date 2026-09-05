namespace OfficeIMO.Workflows;

public sealed partial class OfficeWorkflowRunner {
    private static async Task<string> PublishAsync(
        string stagingPath,
        string requestedPath,
        OfficeWorkflowConflictPolicy policy,
        IOfficeWorkflowPublicationGuard? guard,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        switch (policy) {
            case OfficeWorkflowConflictPolicy.Fail:
                await EnsurePublicationAllowedAsync(guard, requestedPath, false, cancellationToken).ConfigureAwait(false);
                cancellationToken.ThrowIfCancellationRequested();
                File.Move(stagingPath, requestedPath, overwrite: false);
                return requestedPath;
            case OfficeWorkflowConflictPolicy.Replace:
                await EnsurePublicationAllowedAsync(guard, requestedPath, false, cancellationToken).ConfigureAwait(false);
                cancellationToken.ThrowIfCancellationRequested();
                File.Move(stagingPath, requestedPath, overwrite: true);
                return requestedPath;
            case OfficeWorkflowConflictPolicy.Rename:
                for (int suffix = 0; suffix < 10_000; suffix++) {
                    cancellationToken.ThrowIfCancellationRequested();
                    string candidate = suffix == 0 ? requestedPath : AddSuffix(requestedPath, suffix);
                    if (!await CanPublishAsync(guard, candidate, false, cancellationToken).ConfigureAwait(false)) continue;
                    try {
                        File.Move(stagingPath, candidate, overwrite: false);
                        return candidate;
                    } catch (IOException) when (File.Exists(candidate) || Directory.Exists(candidate)) {
                        // Another request owns this candidate. Try the next deterministic suffix.
                    }
                }
                throw new IOException("No available numbered output path could be reserved.");
            default:
                throw new ArgumentOutOfRangeException(nameof(policy), policy, "Unsupported conflict policy.");
        }
    }
    private static async ValueTask<bool> CanPublishAsync(
        IOfficeWorkflowPublicationGuard? guard, string path, bool isDirectory, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        bool allowed = guard is null || await guard.CanPublishAsync(path, isDirectory, cancellationToken).ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
        return allowed;
    }

    private static async ValueTask EnsurePublicationAllowedAsync(
        IOfficeWorkflowPublicationGuard? guard, string path, bool isDirectory, CancellationToken cancellationToken) {
        if (!await CanPublishAsync(guard, path, isDirectory, cancellationToken).ConfigureAwait(false)) {
            throw new IOException("The output destination is in use by the application: " + path);
        }
    }
}
