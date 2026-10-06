using OfficeIMO.Internal;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Features.Workspace;

internal sealed partial class PdfWorkspace {
    private sealed class WorkspaceOutputProgress(IProgress<PdfWorkspaceProgress>? progress) : IProgress<OfficeWorkflowProgress> {
        public void Report(OfficeWorkflowProgress value) => progress?.Report(new(value.Message, value.Fraction));
    }

    private sealed class WorkspaceOutputPublicationGuard(PdfWorkspace workspace, string source, long revision,
        IOfficeWorkflowPublicationGuard? host) : IOfficeWorkflowPublicationGuard {
        public async ValueTask<bool> CanPublishAsync(string path, bool isDirectory, CancellationToken token) {
            token.ThrowIfCancellationRequested();
            if (!IsCurrentAndSeparate(path, isDirectory)) return false;
            if (host is not null && !await host.CanPublishAsync(path, isDirectory, token).ConfigureAwait(false)) return false;
            if (workspace._canPublishOutput is not null && !await workspace._canPublishOutput(path, token).ConfigureAwait(false))
                throw new IOException("The output is already owned by an open document. Choose a different destination.");
            token.ThrowIfCancellationRequested();
            return IsCurrentAndSeparate(path, isDirectory);
        }
        private bool IsCurrentAndSeparate(string path, bool isDirectory) {
            workspace.ThrowIfDisposed();
            if (workspace.Revision != revision || !PathsEqual(source, workspace.Path))
                throw new IOException("The document changed while preparing output.");
            if (OfficeStorageIdentity.AreEquivalent(source, path)) return false;
            if (isDirectory && OfficeStorageIdentity.GetLocalPath(source) is { } localSource &&
                OfficeStorageIdentity.GetLocalPath(path) is { } localOutput && OfficePathIdentity.IsSameOrDescendant(localSource, localOutput)) return false;
            workspace._storage.EnsureWritableLocation(path);
            return true;
        }
    }
}
