namespace OfficeIMO.Studio.Features.Workspace;

internal sealed partial class PdfWorkspace {
    private readonly Dictionary<string, long> _preparedExternalSaves = new(StringComparer.OrdinalIgnoreCase);
    private string? _externalBaseFingerprint;
    internal bool UsesExternalPublication { get; private set; }
    private string RecoveryBaseFingerprint => _externalBaseFingerprint ?? _baseFingerprint;

    private Task<string?> WriteRecoveryAsync(byte[] bytes, long revision, CancellationToken token) =>
        _recoveryStore.WriteAsync(RecoveryIdentity, RecoveryBaseFingerprint, bytes, revision, token, _preparedExternalSaves.Keys.ToArray());

    /// <summary>Captures immutable bytes for a host-owned save, retaining recovery until publication is observed.</summary>
    internal async Task<byte[]> PrepareExternalSaveAsync(CancellationToken token) {
        await _operationGate.WaitAsync(token).ConfigureAwait(false);
        try {
            ThrowIfDisposed();
            byte[] snapshot = (byte[])_bytes.Clone();
            string fingerprint = PdfWorkspaceRecoveryStore.Fingerprint(snapshot);
            bool existed = _preparedExternalSaves.TryGetValue(fingerprint, out long previousRevision);
            _preparedExternalSaves[fingerprint] = _revision;
            try {
                // A crash after the host publishes, but before it acknowledges, must still match the saved snapshot.
                await WriteRecoveryAsync(_bytes, _revision, token).ConfigureAwait(false);
            } catch {
                if (existed) _preparedExternalSaves[fingerprint] = previousRevision;
                else _preparedExternalSaves.Remove(fingerprint);
                throw;
            }
            UsesExternalPublication = true;
            return snapshot;
        } finally { _operationGate.Release(); }
    }

    /// <summary>Acknowledges bytes the host has actually observed at its original document, including an older in-flight save.</summary>
    internal async Task AcknowledgeExternalSaveAsync(string fingerprint, CancellationToken token, string? recoveryIdentity = null) {
        await _operationGate.WaitAsync(token).ConfigureAwait(false);
        try {
            ThrowIfDisposed();
            if (recoveryIdentity is not null && OfficeIMO.Internal.OfficeStorageIdentity.Normalize(recoveryIdentity) != RecoveryIdentity) return;
            if (!_preparedExternalSaves.TryGetValue(fingerprint, out long savedRevision)) return;
            var remaining = _preparedExternalSaves.Where(pair => !string.Equals(pair.Key, fingerprint, StringComparison.OrdinalIgnoreCase)).ToArray();
            if (_revision != savedRevision || remaining.Length > 0) {
                await _recoveryStore.WriteAsync(RecoveryIdentity, fingerprint, _bytes, _revision, token,
                    remaining.Select(pair => pair.Key).ToArray()).ConfigureAwait(false);
            }
            _externalBaseFingerprint = fingerprint;
            _savedRevision = savedRevision;
            _preparedExternalSaves.Remove(fingerprint);
            if (_revision == savedRevision && remaining.Length == 0)
                await _recoveryStore.DeleteAsync(RecoveryIdentity).ConfigureAwait(false);
            RecoveryPath = null;
            Changed?.Invoke(this, EventArgs.Empty);
        } finally { _operationGate.Release(); }
    }
}
