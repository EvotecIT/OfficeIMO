using OfficeIMO.Core.Internal;
namespace OfficeIMO;

internal static class OfficeFilePublisher {
    internal static async Task<string> PublishAsync(
        string stagingPath,
        string requestedPath,
        OfficePublicationPolicy policy,
        IOfficePublicationGuard? guard,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        switch (policy) {
            case OfficePublicationPolicy.Fail:
                await EnsurePublicationAllowedAsync(guard, requestedPath, false, cancellationToken).ConfigureAwait(false);
                cancellationToken.ThrowIfCancellationRequested();
                File.Move(stagingPath, requestedPath);
                return requestedPath;
            case OfficePublicationPolicy.Replace:
                await EnsurePublicationAllowedAsync(guard, requestedPath, false, cancellationToken).ConfigureAwait(false);
                cancellationToken.ThrowIfCancellationRequested();
                OfficeFileCommit.CommitTemporaryFileAtomically(stagingPath, requestedPath);
                return requestedPath;
            case OfficePublicationPolicy.Rename:
                for (int suffix = 0; suffix < 10_000; suffix++) {
                    cancellationToken.ThrowIfCancellationRequested();
                    string candidate = suffix == 0 ? requestedPath : AddSuffix(requestedPath, suffix);
                    if (!await CanPublishAsync(guard, candidate, false, cancellationToken).ConfigureAwait(false)) continue;
                    try {
                        File.Move(stagingPath, candidate);
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
    private static string AddSuffix(string path, int suffix) => Path.Combine(Path.GetDirectoryName(path)!,
        Path.GetFileNameWithoutExtension(path) + " (" + suffix.ToString(System.Globalization.CultureInfo.InvariantCulture) + ")" + Path.GetExtension(path));

    private static async Task<bool> CanPublishAsync(
        IOfficePublicationGuard? guard, string path, bool isDirectory, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        bool allowed = guard is null || await guard.CanPublishAsync(path, isDirectory, cancellationToken).ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
        return allowed;
    }

    private static async Task EnsurePublicationAllowedAsync(
        IOfficePublicationGuard? guard, string path, bool isDirectory, CancellationToken cancellationToken) {
        if (!await CanPublishAsync(guard, path, isDirectory, cancellationToken).ConfigureAwait(false)) {
            throw new IOException("The output destination is in use by the application: " + path);
        }
    }
}
