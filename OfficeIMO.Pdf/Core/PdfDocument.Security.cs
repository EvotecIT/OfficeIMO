namespace OfficeIMO.Pdf;

public sealed partial class PdfDocument {
    /// <summary>Encrypts this unencrypted PDF and returns the rewritten document with preservation proof.</summary>
    public PdfSecurityMutationResult Encrypt(PdfStandardEncryptionOptions encryption) {
        Guard.NotNull(encryption, nameof(encryption));
        return PdfSecurityEditor.Encrypt(GetBytesForOperation(), encryption);
    }

    /// <summary>Attempts to encrypt this unencrypted PDF through the shared mutation planner.</summary>
    public PdfOperationResult<PdfSecurityMutationResult> TryEncrypt(PdfStandardEncryptionOptions encryption) {
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
    public PdfSecurityMutationResult Decrypt(string ownerPassword) {
        Guard.NotNull(ownerPassword, nameof(ownerPassword));
        return PdfSecurityEditor.Decrypt(GetBytesForOperation(), ownerPassword);
    }

    /// <summary>Attempts to remove Standard password security using the current owner password.</summary>
    public PdfOperationResult<PdfSecurityMutationResult> TryDecrypt(string ownerPassword) {
        Guard.NotNull(ownerPassword, nameof(ownerPassword));
        var readOptions = new PdfReadOptions { Password = ownerPassword };
        return TryMutationOperation(
            "Decrypt document",
            PdfPreflightCapability.ManipulatePages,
            PdfMutationOperation.ChangeEncryption,
            _ => Decrypt(ownerPassword),
            options: readOptions,
            executionPreference: PdfMutationExecutionPreference.RequireFullRewrite);
    }

    /// <summary>Replaces Standard password security using the current owner password and returns preservation proof.</summary>
    public PdfSecurityMutationResult Reencrypt(
        string currentOwnerPassword,
        PdfStandardEncryptionOptions newEncryption) {
        Guard.NotNull(currentOwnerPassword, nameof(currentOwnerPassword));
        Guard.NotNull(newEncryption, nameof(newEncryption));
        return PdfSecurityEditor.Reencrypt(GetBytesForOperation(), currentOwnerPassword, newEncryption);
    }

    /// <summary>Attempts to replace Standard password security using the current owner password.</summary>
    public PdfOperationResult<PdfSecurityMutationResult> TryReencrypt(
        string currentOwnerPassword,
        PdfStandardEncryptionOptions newEncryption) {
        Guard.NotNull(currentOwnerPassword, nameof(currentOwnerPassword));
        Guard.NotNull(newEncryption, nameof(newEncryption));
        var readOptions = new PdfReadOptions { Password = currentOwnerPassword };
        return TryMutationOperation(
            "Re-encrypt document",
            PdfPreflightCapability.ManipulatePages,
            PdfMutationOperation.ChangeEncryption,
            _ => Reencrypt(currentOwnerPassword, newEncryption),
            options: readOptions,
            executionPreference: PdfMutationExecutionPreference.RequireFullRewrite);
    }
}
