namespace OfficeIMO.Studio.Native;

internal sealed partial class StudioNativeDocument {
    internal async Task<byte[]> PrepareSaveAsync(CancellationToken token) {
        await EnterAsync(token).ConfigureAwait(false);
        try { return await _workspace.PrepareExternalSaveAsync(token).ConfigureAwait(false); }
        finally { _gate.Release(); }
    }

    internal Task AcknowledgeSaveAsync(string fingerprint, CancellationToken token, string? identity = null) =>
        RunAsync(() => _workspace.AcknowledgeExternalSaveAsync(fingerprint, token, identity), token);

    internal Task SetRecoveryIdentityAsync(string? identity, CancellationToken token) =>
        RunAsync(() => _workspace.SetRecoveryIdentityAsync(identity, token), token);

    internal Task RestoreRecoveryAsync(CancellationToken token) =>
        RunAsync(() => _workspace.RestoreRecoveryAsync(token), token);

    internal Task DiscardRecoveryAsync(CancellationToken token) =>
        RunAsync(() => _workspace.DiscardRecoveryAsync(token), token);
}
