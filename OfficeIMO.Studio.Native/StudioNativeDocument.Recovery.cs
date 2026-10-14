namespace OfficeIMO.Studio.Native;

internal sealed partial class StudioNativeDocument {
    internal Task SetRecoveryIdentityAsync(string? identity, CancellationToken token) =>
        RunAsync(() => _workspace.SetRecoveryIdentityAsync(identity, token), token);

    internal Task RestoreRecoveryAsync(CancellationToken token) =>
        RunAsync(() => _workspace.RestoreRecoveryAsync(token), token);

    internal Task DiscardRecoveryAsync(CancellationToken token) =>
        RunAsync(() => _workspace.DiscardRecoveryAsync(token), token);
}
