using System.Diagnostics;

namespace OfficeIMO.Studio.Features.Workspace;

internal sealed partial class PdfWorkspaceRecoveryStore {
    internal const string LockFileName = ".recovery.lock";

    private FileStream? TryAcquireExclusive() {
        try {
            var options = new FileStreamOptions {
                Mode = FileMode.OpenOrCreate,
                Access = FileAccess.ReadWrite,
                Share = FileShare.None,
                BufferSize = 1,
                Options = FileOptions.Asynchronous
            };
            if (!OperatingSystem.IsWindows()) options.UnixCreateMode = UnixFileMode.UserRead | UnixFileMode.UserWrite;
            return new FileStream(Path.Combine(_root, LockFileName), options);
        } catch (IOException) {
            return null;
        } catch (UnauthorizedAccessException) {
            return null;
        }
    }

    private async Task<FileStream> AcquireExclusiveAsync(CancellationToken token) {
        token.ThrowIfCancellationRequested();
        Directory.CreateDirectory(_root);
        var elapsed = Stopwatch.StartNew();
        while (true) {
            token.ThrowIfCancellationRequested();
            FileStream? lease = TryAcquireExclusive();
            if (lease is not null) return lease;
            if (elapsed.Elapsed >= TimeSpan.FromSeconds(10)) {
                throw new IOException("Recovery storage is busy or unavailable. Please retry the operation.");
            }
            await Task.Delay(25, token).ConfigureAwait(false);
        }
    }
}
