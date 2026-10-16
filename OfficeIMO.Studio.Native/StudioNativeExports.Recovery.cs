using System.Runtime.InteropServices;

namespace OfficeIMO.Studio.Native;

internal static unsafe partial class StudioNativeExports {
    /// <summary>Returns immutable PDF bytes for a host-owned save. Preparation does not acknowledge publication.</summary>
    [UnmanagedCallersOnly(EntryPoint = "officeimo_document_prepare_save")]
    internal static int DocumentPrepareSave(nint document, nint cancellation, byte** pdf, nuint* length, byte** error) {
        if (pdf is null || length is null) return Fail(error, new ArgumentNullException(pdf is null ? nameof(pdf) : nameof(length)));
        *pdf = null; *length = 0;
        return WithDocument(document, error, cancellation, (target, token) => {
            byte[] bytes = target.PrepareSaveAsync(token).GetAwaiter().GetResult();
            *pdf = StudioNativeAbi.ToNative(bytes);
            *length = (nuint)bytes.Length;
        });
    }

    /// <summary>Acknowledges the SHA-256 of a prepared snapshot observed at the host's saved document.</summary>
    [UnmanagedCallersOnly(EntryPoint = "officeimo_document_acknowledge_save")]
    internal static int DocumentAcknowledgeSave(nint document, byte* fingerprint, byte* identity, nint cancellation, byte** error) {
        return AcknowledgeSave(document, fingerprint, identity, cancellation, error, supersedesEarlierPreparations: false);
    }

    /// <summary>Acknowledges ordered publication, retaining recovery when identical bytes leave its occurrence ambiguous.</summary>
    [UnmanagedCallersOnly(EntryPoint = "officeimo_document_acknowledge_ordered_save")]
    internal static int DocumentAcknowledgeOrderedSave(nint document, byte* fingerprint, byte* identity, nint cancellation, byte** error) =>
        AcknowledgeSave(document, fingerprint, identity, cancellation, error, supersedesEarlierPreparations: true);

    private static int AcknowledgeSave(nint document, byte* fingerprint, byte* identity, nint cancellation, byte** error,
        bool supersedesEarlierPreparations) {
        string? value = StudioNativeAbi.FromUtf8(fingerprint);
        string? source = StudioNativeAbi.FromUtf8(identity);
        return WithDocument(document, error, cancellation, (target, token) =>
            target.AcknowledgeSaveAsync(value ?? throw new ArgumentNullException(nameof(fingerprint)), token,
                source ?? throw new ArgumentNullException(nameof(identity)), supersedesEarlierPreparations).GetAwaiter().GetResult());
    }

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
