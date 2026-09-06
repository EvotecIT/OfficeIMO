using OfficeIMO.Core.Internal;

namespace OfficeIMO.Studio.Features.Workspace;

internal sealed partial class PdfWorkspace {
    private async Task VerifyOutputDestinationAsync(string destination, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        ValidateExportDestination(destination);
        if (_canPublishOutput is not null && !await _canPublishOutput(destination, cancellationToken).ConfigureAwait(false)) {
            throw new IOException("The output is already owned by an open document. Choose a different destination.");
        }
        cancellationToken.ThrowIfCancellationRequested();
    }

    private async Task WriteWorkspaceOutputAsync(string destination, Func<Stream, CancellationToken, Task> writer,
        CancellationToken cancellationToken) {
        await VerifyOutputDestinationAsync(destination, cancellationToken).ConfigureAwait(false);
        await OfficeFileCommit.WriteAsync(destination, async (stream, token) => {
            await writer(stream, token).ConfigureAwait(false);
            await VerifyOutputDestinationAsync(destination, token).ConfigureAwait(false);
        }, cancellationToken: cancellationToken).ConfigureAwait(false);
    }
}
