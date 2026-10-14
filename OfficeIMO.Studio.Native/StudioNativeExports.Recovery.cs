using System.Runtime.InteropServices;

namespace OfficeIMO.Studio.Native;

internal static unsafe partial class StudioNativeExports {
    /// <summary>Opens a private working file while recovery follows the host's stable original-document identity.</summary>
    [UnmanagedCallersOnly(EntryPoint = "officeimo_document_open_with_recovery_identity")]
    internal static int DocumentOpenWithRecoveryIdentity(byte* path, byte* password, byte* identity,
        nint cancellation, nint* document, byte** error) => OpenDocument(path, password, identity, cancellation, document, error);

    [UnmanagedCallersOnly(EntryPoint = "officeimo_document_set_recovery_identity")]
    internal static int DocumentSetRecoveryIdentity(nint document, byte* identity, nint cancellation, byte** error) {
        string? value = StudioNativeAbi.FromUtf8(identity);
        return WithDocument(document, error, cancellation, (target, token) => target.SetRecoveryIdentityAsync(value, token).GetAwaiter().GetResult());
    }

    [UnmanagedCallersOnly(EntryPoint = "officeimo_document_restore_recovery")]
    internal static int DocumentRestoreRecovery(nint document, nint cancellation, byte** error) =>
        WithDocument(document, error, cancellation, (target, token) => target.RestoreRecoveryAsync(token).GetAwaiter().GetResult());

    [UnmanagedCallersOnly(EntryPoint = "officeimo_document_discard_recovery")]
    internal static int DocumentDiscardRecovery(nint document, nint cancellation, byte** error) =>
        WithDocument(document, error, cancellation, (target, token) => target.DiscardRecoveryAsync(token).GetAwaiter().GetResult());
}
