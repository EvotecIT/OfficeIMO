namespace OfficeIMO.Studio.Features.Workspace;

internal sealed partial class PdfWorkspace {
    // Repeated bytes retain a conservative interval, without allocating one record per duplicate preparation.
    private sealed record ExternalSavePreparation(long Revision, long FirstPossibleOrder, long LastOrder);
    private readonly Dictionary<string, ExternalSavePreparation> _preparedExternalSaves = new(StringComparer.OrdinalIgnoreCase);
    private long _externalSavePreparationOrder;
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
            bool existed = _preparedExternalSaves.TryGetValue(fingerprint, out var previousPreparation);
            long order = checked(++_externalSavePreparationOrder);
            _preparedExternalSaves[fingerprint] = new(_revision, previousPreparation?.FirstPossibleOrder ?? order, order);
            try {
                // A crash after the host publishes, but before it acknowledges, must still match the saved snapshot.
                await WriteRecoveryAsync(_bytes, _revision, token).ConfigureAwait(false);
            } catch {
                if (existed) _preparedExternalSaves[fingerprint] = previousPreparation!;
                else _preparedExternalSaves.Remove(fingerprint);
                throw;
            }
            UsesExternalPublication = true;
            return snapshot;
        } finally { _operationGate.Release(); }
    }

    /// <summary>Acknowledges bytes the host has actually observed at its original document, including an older in-flight save.</summary>
    /// <param name="supersedesEarlierPreparations">Use only when the host serializes publication and the observed
    /// snapshot supersedes earlier preparations. Identical-byte ambiguity retains recovery until distinct readback.
    /// General hosts may still publish older in-flight snapshots.</param>
    internal async Task AcknowledgeExternalSaveAsync(string fingerprint, CancellationToken token, string? recoveryIdentity = null,
        bool supersedesEarlierPreparations = false) {
        await _operationGate.WaitAsync(token).ConfigureAwait(false);
        try {
            ThrowIfDisposed();
            if (recoveryIdentity is not null && OfficeIMO.Internal.OfficeStorageIdentity.Normalize(recoveryIdentity) != RecoveryIdentity) return;
            if (!_preparedExternalSaves.TryGetValue(fingerprint, out var observed)) return;
            long savedRevision = observed.Revision;
            bool hasOtherBytes = _preparedExternalSaves.Keys.Any(key => !string.Equals(key, fingerprint, StringComparison.OrdinalIgnoreCase));
            // Readback cannot distinguish A1 from identical A3, nor a duplicate notification from a new publication.
            // Retain ambiguous in-flight bytes until a distinct observation advances the publication frontier.
            if (supersedesEarlierPreparations && hasOtherBytes &&
                string.Equals(fingerprint, RecoveryBaseFingerprint, StringComparison.OrdinalIgnoreCase)) return;
            var remaining = supersedesEarlierPreparations
                ? (hasOtherBytes ? _preparedExternalSaves.Where(pair => pair.Value.LastOrder > observed.FirstPossibleOrder)
                    .Select(pair => new KeyValuePair<string, ExternalSavePreparation>(pair.Key,
                        pair.Value with { FirstPossibleOrder = Math.Max(pair.Value.FirstPossibleOrder, observed.FirstPossibleOrder + 1) })).ToArray() : [])
                : _preparedExternalSaves.Where(pair => !string.Equals(pair.Key, fingerprint, StringComparison.OrdinalIgnoreCase)).ToArray();
            if (_revision != savedRevision || remaining.Length > 0) {
                await _recoveryStore.WriteAsync(RecoveryIdentity, fingerprint, _bytes, _revision, token,
                    remaining.Select(pair => pair.Key).ToArray()).ConfigureAwait(false);
            } else {
                // Commit recovery I/O before advancing the in-memory saved state.
                token.ThrowIfCancellationRequested();
                await _recoveryStore.DeleteAsync(RecoveryIdentity).ConfigureAwait(false);
            }
            _externalBaseFingerprint = fingerprint;
            _savedRevision = savedRevision;
            _preparedExternalSaves.Clear();
            foreach (var retained in remaining) _preparedExternalSaves.Add(retained.Key, retained.Value);
            RecoveryPath = null;
            Changed?.Invoke(this, EventArgs.Empty);
        } finally { _operationGate.Release(); }
    }
}
