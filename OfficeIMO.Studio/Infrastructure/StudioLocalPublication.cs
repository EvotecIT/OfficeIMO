using Avalonia.Threading;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Studio.Infrastructure;

/// <summary>Stages local output before checking live UI ownership and committing without another UI yield.</summary>
internal static class StudioLocalPublication {
    internal static async Task WriteAsync(string path, byte[] bytes, Action authorize, CancellationToken token) {
        string? staging = null;
        try {
            staging = await OfficeFileCommit.StageAllBytesAsync(path, bytes, token).ConfigureAwait(false);
            await Dispatcher.UIThread.InvokeAsync(() => {
                token.ThrowIfCancellationRequested();
                authorize();
                OfficeFileCommit.CommitTemporaryFileAtomically(staging, path);
                staging = null;
            }, DispatcherPriority.Normal, token);
        } finally {
            OfficeFileCommit.DeleteIfExists(staging);
        }
    }
}
