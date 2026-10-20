using System.Threading;

namespace OfficeIMO.Pdf;

public sealed partial class PdfDocument {
    /// <summary>Encrypts this unencrypted PDF and returns the rewritten document with preservation proof.</summary>
    internal PdfSecurityMutationResult Encrypt(PdfStandardEncryptionOptions encryption, CancellationToken cancellationToken = default) {
        Guard.NotNull(encryption, nameof(encryption));
        return PdfSecurityEditor.Encrypt(GetBytesForOperation(cancellationToken), encryption, ReadOptions, cancellationToken: cancellationToken);
    }

    /// <summary>Attempts to encrypt this unencrypted PDF through the shared mutation planner.</summary>
    internal PdfOperationResult<PdfSecurityMutationResult> EncryptResult(PdfStandardEncryptionOptions encryption, CancellationToken cancellationToken = default) {
        Guard.NotNull(encryption, nameof(encryption));
        return TryMutationOperation(
            "Encrypt document",
            PdfPreflightCapability.ManipulatePages,
            PdfMutationOperation.ChangeEncryption,
            _ => Encrypt(encryption, cancellationToken),
            options: ReadOptions,
            executionPreference: PdfMutationExecutionPreference.RequireFullRewrite, cancellationToken: cancellationToken);
    }

    /// <summary>Removes Standard password security using the current owner password and returns preservation proof.</summary>
    internal PdfSecurityMutationResult Decrypt(string ownerPassword, CancellationToken cancellationToken = default) {
        Guard.NotNull(ownerPassword, nameof(ownerPassword));
        return PdfSecurityEditor.Decrypt(GetBytesForOperation(cancellationToken), ownerPassword, ReadOptions, cancellationToken: cancellationToken);
    }

    /// <summary>Attempts to remove Standard password security using the current owner password.</summary>
    internal PdfOperationResult<PdfSecurityMutationResult> DecryptResult(string ownerPassword, CancellationToken cancellationToken = default) {
        Guard.NotNull(ownerPassword, nameof(ownerPassword));
        PdfLoadOptions readOptions = PdfLoadOptions.WithPassword(ReadOptions, ownerPassword);
        return TryMutationOperation(
            "Decrypt document",
            PdfPreflightCapability.ManipulatePages,
            PdfMutationOperation.ChangeEncryption,
            _ => Decrypt(ownerPassword, cancellationToken),
            options: readOptions,
            executionPreference: PdfMutationExecutionPreference.RequireFullRewrite, cancellationToken: cancellationToken);
    }

    /// <summary>Replaces Standard password security using the current owner password and returns preservation proof.</summary>
    internal PdfSecurityMutationResult Reencrypt(
        string currentOwnerPassword,
        PdfStandardEncryptionOptions newEncryption,
        CancellationToken cancellationToken = default) {
        Guard.NotNull(currentOwnerPassword, nameof(currentOwnerPassword));
        Guard.NotNull(newEncryption, nameof(newEncryption));
        return PdfSecurityEditor.Reencrypt(GetBytesForOperation(cancellationToken), currentOwnerPassword, newEncryption, ReadOptions, cancellationToken: cancellationToken);
    }

    /// <summary>Attempts to replace Standard password security using the current owner password.</summary>
    internal PdfOperationResult<PdfSecurityMutationResult> ReencryptResult(
        string currentOwnerPassword,
        PdfStandardEncryptionOptions newEncryption,
        CancellationToken cancellationToken = default) {
        Guard.NotNull(currentOwnerPassword, nameof(currentOwnerPassword));
        Guard.NotNull(newEncryption, nameof(newEncryption));
        PdfLoadOptions readOptions = PdfLoadOptions.WithPassword(ReadOptions, currentOwnerPassword);
        return TryMutationOperation(
            "Re-encrypt document",
            PdfPreflightCapability.ManipulatePages,
            PdfMutationOperation.ChangeEncryption,
            _ => Reencrypt(currentOwnerPassword, newEncryption, cancellationToken),
            options: readOptions,
            executionPreference: PdfMutationExecutionPreference.RequireFullRewrite, cancellationToken: cancellationToken);
    }
}
