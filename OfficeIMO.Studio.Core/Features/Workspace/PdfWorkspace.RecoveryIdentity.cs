using OfficeIMO.Internal;

namespace OfficeIMO.Studio.Features.Workspace;

internal sealed partial class PdfWorkspace {
    private string? _recoveryIdentity;
    private string RecoveryIdentity => _recoveryIdentity ?? Path;

    /// <summary>Moves recovery ownership when a host's original document moves, independently of its private working file.</summary>
    internal async Task SetRecoveryIdentityAsync(string? identity, CancellationToken token) {
        string? normalized = identity is null ? null : OfficeStorageIdentity.Normalize(identity);
        if (normalized?.Length > 4096) throw new ArgumentException("The recovery identity is too long.", nameof(identity));
        await _operationGate.WaitAsync(token).ConfigureAwait(false);
        try {
            ThrowIfDisposed();
            string previous = RecoveryIdentity, next = normalized ?? Path;
            if (previous == next) return;
            if (IsDirty) await _recoveryStore.WriteAsync(next, _baseFingerprint, _bytes, _revision, token).ConfigureAwait(false);
            await _recoveryStore.DeleteAsync(previous, token).ConfigureAwait(false);
            _recoveryIdentity = normalized;
            RecoveryPath = IsDirty ? null : _recoveryStore.Find(next, _baseFingerprint);
            Changed?.Invoke(this, EventArgs.Empty);
        } finally { _operationGate.Release(); }
    }
}
