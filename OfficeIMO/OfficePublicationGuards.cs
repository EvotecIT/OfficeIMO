namespace OfficeIMO;

internal static class OfficePublicationGuards {
    internal static IOfficePublicationGuard Combine(IOfficePublicationGuard? host, IOfficePublicationGuard source) =>
        host == null ? source : new Combined(host, source);

    private sealed class Combined(IOfficePublicationGuard host, IOfficePublicationGuard source) : IOfficePublicationGuard, IOfficeStagingGuard {
        public async Task<bool> CanPublishAsync(string path, bool isDirectory, CancellationToken cancellationToken) =>
            await host.CanPublishAsync(path, isDirectory, cancellationToken).ConfigureAwait(false)
            && await source.CanPublishAsync(path, isDirectory, cancellationToken).ConfigureAwait(false);
        public async Task EnsureStagingDirectoryAllowedAsync(string directory, CancellationToken cancellationToken) {
            if (host is IOfficeStagingGuard hostStaging) await hostStaging.EnsureStagingDirectoryAllowedAsync(directory, cancellationToken).ConfigureAwait(false);
            if (source is IOfficeStagingGuard sourceStaging) await sourceStaging.EnsureStagingDirectoryAllowedAsync(directory, cancellationToken).ConfigureAwait(false);
        }
    }
}
