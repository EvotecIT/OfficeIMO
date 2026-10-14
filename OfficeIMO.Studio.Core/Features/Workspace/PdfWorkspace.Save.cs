using System.Security.Cryptography;
using OfficeIMO.Pdf;
using OfficeIMO.Core.Internal;
using OfficeIMO.Internal;
using OfficeIMO.Studio.Infrastructure;

namespace OfficeIMO.Studio.Features.Workspace;

internal sealed partial class PdfWorkspace {
    /// <summary>
    /// Writes the open revision to another file, as a share or an export does; the document keeps its own location and
    /// saved state.
    /// </summary>
    internal async Task SaveCopyAsync(string path, CancellationToken cancellationToken) {
        ThrowIfDisposed();
        string destination = ValidateCopyDestination(path);
        await _operationGate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try {
            byte[] bytes = _bytes;
            await WriteWorkspaceOutputAsync(destination,
                async (stream, token) => await LoadDocument(bytes).SaveAsync(stream, token).ConfigureAwait(false),
                cancellationToken).ConfigureAwait(false);
        } finally {
            _operationGate.Release();
        }
    }

    /// <summary>Uses the workspace's storage identity and write policy for any operation that promises a separate copy.</summary>
    internal string ValidateCopyDestination(string path) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("A destination is required.", nameof(path));
        string destination = OfficeStorageIdentity.Normalize(path);
        if (OfficeStorageIdentity.AreEquivalent(destination, Path)) {
            throw new ArgumentException("A copy can't replace the document it was made from.", nameof(path));
        }
        _storage.EnsureWritableLocation(destination);
        return destination;
    }

    internal async Task SaveAsync(string? path, CancellationToken cancellationToken, IProgress<PdfWorkspaceProgress>? progress = null) {
        ThrowIfDisposed();
        string destination = string.IsNullOrWhiteSpace(path) ? Path : OfficeStorageIdentity.Normalize(path);
        _storage.EnsureWritableLocation(destination);
        await _operationGate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try {
            string previousPath = Path;
            string previousRecoveryIdentity = RecoveryIdentity;
            progress?.Report(new PdfWorkspaceProgress("Saving PDF", 0.2D));
            PdfSaveResult? saved = null;
            string? savedSourceIdentity = null;
            if (_storage.UsesProviderPublication(destination)) {
                bool replacingSource = OfficeStorageIdentity.AreEquivalent(destination, Path);
                if (!replacingSource) await VerifyOutputDestinationAsync(destination, cancellationToken).ConfigureAwait(false);
                using var serialized = new OfficeBoundedMemoryStream(StudioDocumentStorage.MaximumDocumentBytes);
                saved = await LoadDocument(_bytes).SaveAsync(serialized, cancellationToken).ConfigureAwait(false);
                if (replacingSource) {
                    await _recoveryStore.WriteAsync(RecoveryIdentity, _baseFingerprint, _bytes, _revision, cancellationToken).ConfigureAwait(false);
                }
                try {
                    StudioStoragePublication publication = await _storage.PublishAsync(destination, serialized.ToArray(), replacingSource ? _baseFingerprint : null,
                        async token => {
                            if (!replacingSource) await VerifyOutputDestinationAsync(destination, token).ConfigureAwait(false);
                            else if (await _storage.ReadIdentityAsync(Path, token).ConfigureAwait(false) != _sourceIdentityKey) {
                                throw new IOException("The source was replaced after it was opened. Save to a different destination.");
                            }
                        },
                        cancellationToken).ConfigureAwait(false);
                    savedSourceIdentity = publication.Identity;
                } catch (Exception error) when (replacingSource && OfficeStreamPublication.MayHaveChangedDestination(error)) {
                    // No historical revision is known to match a partially written provider destination.
                    // Undo must not make this document silently closeable without preserving a copy.
                    _savedRevision = -1;
                    Changed?.Invoke(this, EventArgs.Empty);
                    throw;
                }
            } else if (OfficeStorageIdentity.AreEquivalent(destination, Path)) {
                if (OfficePathIdentity.GetPhysicalIdentityKey(Path) != _sourceIdentityKey) {
                    throw new IOException("The source PDF was replaced or moved after it was opened. Use Save As to preserve your edits in a different file.");
                }
                // Publish to the resolved file while preserving the user's symlink itself.
                destination = OfficePathIdentity.ResolvePhysicalPath(destination);
                await OfficeFileCommit.WriteIfUnchangedAsync(destination,
                    async (stream, token) => saved = await LoadDocument(_bytes).SaveAsync(stream, token).ConfigureAwait(false),
                    candidate => {
                        using var stream = new FileStream(candidate, FileMode.Open, FileAccess.Read, FileShare.Read | FileShare.Delete);
                        if (OfficePathIdentity.GetPhysicalIdentityKey(candidate, stream.SafeFileHandle) != _sourceIdentityKey) return false;
                        return string.Equals(_baseFingerprint, Convert.ToHexString(SHA256.HashData(stream)), StringComparison.OrdinalIgnoreCase);
                    }, cancellationToken).ConfigureAwait(false);
            } else {
                await WriteWorkspaceOutputAsync(destination,
                    async (stream, token) => saved = await LoadDocument(_bytes).SaveAsync(stream, token).ConfigureAwait(false),
                    cancellationToken).ConfigureAwait(false);
            }
            Path = destination;
            _sourceIdentityKey = savedSourceIdentity ?? OfficePathIdentity.GetPhysicalIdentityKey(destination);
            _baseFingerprint = saved!.Pipeline.Output!.Sha256.ToUpperInvariant();
            _savedRevision = _revision;
            // Publication has completed; cleanup must not be interrupted by late cancellation.
            try {
                await _recoveryStore.DeleteAsync(previousRecoveryIdentity).ConfigureAwait(false);
                if (previousRecoveryIdentity != RecoveryIdentity) await _recoveryStore.DeleteAsync(RecoveryIdentity).ConfigureAwait(false);
            } catch (Exception error) when (error is IOException or UnauthorizedAccessException) {
                Changed?.Invoke(this, EventArgs.Empty);
                throw new IOException("The PDF was saved, but its stored recovery data could not be removed. Clear stored recovery data in Settings when storage is available.", error);
            }
            RecoveryPath = null;
            progress?.Report(new PdfWorkspaceProgress("Saved", 1D));
            Changed?.Invoke(this, EventArgs.Empty);
        } finally {
            _operationGate.Release();
        }
    }

}
