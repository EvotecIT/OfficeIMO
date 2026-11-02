using OfficeIMO.Internal;

namespace OfficeIMO.Workflows;

public sealed partial class OfficeWorkflowRunner {
    /// <summary>Checks actual provider locations again after deferred destination creation and host authorization.</summary>
    private sealed class DistinctWorkflowOutputPublicationGuard(IOfficePublicationGuard? host,
        Func<IEnumerable<string>> publishedLocations) : IOfficePublicationGuard {
        public async Task<bool> CanPublishAsync(string path, bool isDirectory, CancellationToken token) {
            if (publishedLocations().Any(previous => OfficeStorageIdentity.AreEquivalent(previous, path))) return false;
            if (host is not null && !await host.CanPublishAsync(path, isDirectory, token).ConfigureAwait(false)) return false;
            token.ThrowIfCancellationRequested();
            return !publishedLocations().Any(previous => OfficeStorageIdentity.AreEquivalent(previous, path));
        }
    }
}
