namespace OfficeIMO.Pdf;

public sealed partial class PdfDocument {
    /// <summary>Encrypts this unencrypted PDF and returns the rewritten document with preservation proof.</summary>
    internal PdfSecurityMutationResult Encrypt(PdfStandardEncryptionOptions encryption) {
        Guard.NotNull(encryption, nameof(encryption));
        return PdfSecurityEditor.Encrypt(GetBytesForOperation(), encryption, ReadOptions);
    }

    /// <summary>Attempts to encrypt this unencrypted PDF through the shared mutation planner.</summary>
    internal PdfOperationResult<PdfSecurityMutationResult> TryEncrypt(PdfStandardEncryptionOptions encryption) {
        Guard.NotNull(encryption, nameof(encryption));
        return TryMutationOperation(
            "Encrypt document",
            PdfPreflightCapability.ManipulatePages,
            PdfMutationOperation.ChangeEncryption,
            _ => Encrypt(encryption),
            options: ReadOptions,
            executionPreference: PdfMutationExecutionPreference.RequireFullRewrite);
    }

    /// <summary>Removes Standard password security using the current owner password and returns preservation proof.</summary>
    internal PdfSecurityMutationResult Decrypt(string ownerPassword) {
        Guard.NotNull(ownerPassword, nameof(ownerPassword));
        return PdfSecurityEditor.Decrypt(GetBytesForOperation(), ownerPassword, ReadOptions);
    }

    /// <summary>Attempts to remove Standard password security using the current owner password.</summary>
    internal PdfOperationResult<PdfSecurityMutationResult> TryDecrypt(string ownerPassword) {
        Guard.NotNull(ownerPassword, nameof(ownerPassword));
        PdfLoadOptions readOptions = PdfLoadOptions.WithPassword(ReadOptions, ownerPassword);
        return TryMutationOperation(
            "Decrypt document",
            PdfPreflightCapability.ManipulatePages,
            PdfMutationOperation.ChangeEncryption,
            _ => Decrypt(ownerPassword),
            options: readOptions,
            executionPreference: PdfMutationExecutionPreference.RequireFullRewrite);
    }

    /// <summary>Replaces Standard password security using the current owner password and returns preservation proof.</summary>
    internal PdfSecurityMutationResult Reencrypt(
        string currentOwnerPassword,
        PdfStandardEncryptionOptions newEncryption) {
        Guard.NotNull(currentOwnerPassword, nameof(currentOwnerPassword));
        Guard.NotNull(newEncryption, nameof(newEncryption));
        return PdfSecurityEditor.Reencrypt(GetBytesForOperation(), currentOwnerPassword, newEncryption, ReadOptions);
    }

    /// <summary>Attempts to replace Standard password security using the current owner password.</summary>
    internal PdfOperationResult<PdfSecurityMutationResult> TryReencrypt(
        string currentOwnerPassword,
        PdfStandardEncryptionOptions newEncryption) {
        Guard.NotNull(currentOwnerPassword, nameof(currentOwnerPassword));
        Guard.NotNull(newEncryption, nameof(newEncryption));
        PdfLoadOptions readOptions = PdfLoadOptions.WithPassword(ReadOptions, currentOwnerPassword);
        return TryMutationOperation(
            "Re-encrypt document",
            PdfPreflightCapability.ManipulatePages,
            PdfMutationOperation.ChangeEncryption,
            _ => Reencrypt(currentOwnerPassword, newEncryption),
            options: readOptions,
            executionPreference: PdfMutationExecutionPreference.RequireFullRewrite);
    }
}