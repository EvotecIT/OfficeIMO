using OfficeIMO.Core.Internal;

namespace OfficeIMO.Studio.Features.Workspace;

internal sealed partial class PdfWorkspace {
    private async Task VerifyOutputDestinationAsync(string destination, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        _storage.EnsureWritableLocation(destination);
        ValidateExportDestination(destination);
        if (_canPublishOutput is not null && !await _canPublishOutput(destination, cancellationToken).ConfigureAwait(false)) {
            throw new IOException("The output is already owned by an open document. Choose a different destination.");
        }
        cancellationToken.ThrowIfCancellationRequested();
    }

    private async Task WriteWorkspaceOutputAsync(string destination, Func<Stream, CancellationToken, Task> writer,
        CancellationToken cancellationToken, Func<CancellationToken, Task>? verifyRelatedArtifacts = null) {
        async Task VerifyAsync(CancellationToken token) {
            await VerifyOutputDestinationAsync(destination, token).ConfigureAwait(false);
            if (verifyRelatedArtifacts is not null) await verifyRelatedArtifacts(token).ConfigureAwait(false);
        }
        await VerifyAsync(cancellationToken).ConfigureAwait(false);
        if (_storage.UsesProviderPublication(destination)) {
            using var serialized = new OfficeBoundedMemoryStream(Infrastructure.StudioStorageAccess.MaximumDocumentBytes);
            await writer(serialized, cancellationToken).ConfigureAwait(false);
            await _storage.PublishAsync(destination, serialized.ToArray(), expectedFingerprint: null,
                VerifyAsync, cancellationToken).ConfigureAwait(false);
            return;
        }
        await OfficeFileCommit.WriteAsync(destination, async (stream, token) => {
            await writer(stream, token).ConfigureAwait(false);
            await VerifyAsync(token).ConfigureAwait(false);
        }, cancellationToken: cancellationToken).ConfigureAwait(false);
    }
}
