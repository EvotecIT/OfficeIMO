using Avalonia.Threading;
using OfficeIMO.Core.Internal;
using OfficeIMO.Internal;
using System.Security.Cryptography;
using OfficeIMO.Studio.Infrastructure.Localization;

namespace OfficeIMO.Studio.Infrastructure;

/// <summary>Stages local output before checking live UI ownership and committing without another UI yield.</summary>
internal static class StudioLocalPublication {
    /// <summary>Publishes a reviewed local file while preserving a changed or newly occupied destination.</summary>
    internal static async Task<StudioStoragePublication> WriteVerifiedAsync(string path, byte[] bytes,
        string? expectedFingerprint, Func<CancellationToken, Task> authorize, CancellationToken token, IStudioLocalizer? localizer = null) {
        localizer ??= StudioLocalization.Current;
        if (bytes.LongLength > StudioDocumentStorage.MaximumDocumentBytes)
            throw new IOException(localizer.Get("Output.PublicationTooLarge"));
        Dispatcher uiDispatcher = Dispatcher.UIThread;
        await uiDispatcher.InvokeAsync(() => authorize(token), DispatcherPriority.Normal).ConfigureAwait(false);
        string fingerprint = Convert.ToHexString(SHA256.HashData(bytes));
        string? staging = null;
        try {
            staging = await OfficeFileCommit.StageAllBytesAsync(path, bytes, token).ConfigureAwait(false);
            return await uiDispatcher.InvokeAsync(async () => {
                await authorize(token);
                token.ThrowIfCancellationRequested();
                // The final authorization and commit share one UI continuation. No UI operation
                // can take ownership of the destination between those two steps.
                if (expectedFingerprint is null) {
                    OfficeFileCommit.CommitTemporaryFileAtomically(staging, path, OfficeFileCommit.ConflictPolicy.FailIfExists);
                } else if (!OfficeFileCommit.TryCommitTemporaryFileAtomicallyIfDestinationUnchanged(staging, path,
                    candidate => MatchesFingerprint(candidate, expectedFingerprint),
                    installed => MatchesFingerprint(installed, fingerprint))) {
                    throw new IOException(localizer.Get("Output.DestinationChanged"));
                }
                staging = null;
                using var published = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read | FileShare.Delete);
                string identity = OfficePathIdentity.GetPhysicalIdentityKey(path, published.SafeFileHandle);
                if (published.Length > StudioDocumentStorage.MaximumDocumentBytes ||
                    !fingerprint.Equals(Convert.ToHexString(SHA256.HashData(published)), StringComparison.OrdinalIgnoreCase) ||
                    OfficePathIdentity.GetPhysicalIdentityKey(path) != identity)
                    throw new IOException(localizer.Get("Output.VerificationChanged"));
                return new StudioStoragePublication(fingerprint, identity);
            }, DispatcherPriority.Normal);
        } finally { OfficeFileCommit.DeleteIfExists(staging); }
    }

    private static bool MatchesFingerprint(string path, string expected) {
        try {
            using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read | FileShare.Delete);
            return stream.Length <= StudioDocumentStorage.MaximumDocumentBytes &&
                expected.Equals(Convert.ToHexString(SHA256.HashData(stream)), StringComparison.OrdinalIgnoreCase);
        } catch (FileNotFoundException) { return false; }
    }

    internal static async Task WriteAsync(string path, byte[] bytes, Action authorize, CancellationToken token) {
        Dispatcher uiDispatcher = Dispatcher.UIThread;
        string? staging = null;
        try {
            staging = await OfficeFileCommit.StageAllBytesAsync(path, bytes, token).ConfigureAwait(false);
            await uiDispatcher.InvokeAsync(() => {
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
