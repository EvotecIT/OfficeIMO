using OfficeIMO.Studio.Infrastructure;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Features.Shell;

/// <summary>Keeps application-owned recovery copies outside workflow publication destinations.</summary>
internal sealed class StudioProtectedPublicationGuard(StudioStorageAccess storage, IOfficeWorkflowPublicationGuard? inner)
    : IOfficeWorkflowPublicationGuard {
    public async ValueTask<bool> CanPublishAsync(string path, bool isDirectory, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        if (storage.IsRecoveryLocation(path, isDirectory)) return false;
        return inner is null || await inner.CanPublishAsync(path, isDirectory, token).ConfigureAwait(false);
    }
}
