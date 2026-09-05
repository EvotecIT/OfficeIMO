namespace OfficeIMO.Pdf;

/// <summary>Password encryption and digital-signature operations for one PDF document.</summary>
public sealed class PdfDocumentSecurity {
    private readonly PdfDocument _document;

    internal PdfDocumentSecurity(PdfDocument document) => _document = document;

    /// <summary>Encrypts an unencrypted PDF and returns the rewritten document with preservation proof.</summary>
    public PdfSecurityMutationResult Encrypt(PdfStandardEncryptionOptions encryption) => _document.Encrypt(encryption);

    /// <summary>Attempts to encrypt an unencrypted PDF through the shared mutation planner.</summary>
    public PdfOperationResult<PdfSecurityMutationResult> TryEncrypt(PdfStandardEncryptionOptions encryption) => _document.TryEncrypt(encryption);

    /// <summary>Removes Standard password security using the current owner password.</summary>
    public PdfSecurityMutationResult Decrypt(string ownerPassword) => _document.Decrypt(ownerPassword);

    /// <summary>Attempts to remove Standard password security using the current owner password.</summary>
    public PdfOperationResult<PdfSecurityMutationResult> TryDecrypt(string ownerPassword) => _document.TryDecrypt(ownerPassword);

    /// <summary>Replaces Standard password security using the current owner password.</summary>
    public PdfSecurityMutationResult Reencrypt(string currentOwnerPassword, PdfStandardEncryptionOptions newEncryption) =>
        _document.Reencrypt(currentOwnerPassword, newEncryption);

    /// <summary>Attempts to replace Standard password security using the current owner password.</summary>
    public PdfOperationResult<PdfSecurityMutationResult> TryReencrypt(string currentOwnerPassword, PdfStandardEncryptionOptions newEncryption) =>
        _document.TryReencrypt(currentOwnerPassword, newEncryption);

    /// <summary>Validates signature structure, byte ranges, and preservation markers.</summary>
    public PdfSignatureValidationReport ValidateSignatures(PdfLoadOptions? options = null) => _document.ValidateSignatures(options);

    /// <summary>Validates signatures with caller-provided CMS, trust, timestamp, and revocation policy.</summary>
    public PdfSignatureValidationReport ValidateSignatures(IPdfSignatureCryptographyProvider cryptographyProvider, PdfLoadOptions? options = null) =>
        _document.ValidateSignatures(cryptographyProvider, options);

    /// <summary>Prepares a placeholder and byte ranges for an externally produced signature.</summary>
    public PdfExternalSignaturePreparation PrepareExternalSignature(PdfExternalSignatureOptions? options = null) =>
        _document.PrepareExternalSignature(options);

    /// <summary>Attempts to prepare an external signature through document preflight.</summary>
    public PdfOperationResult<PdfExternalSignaturePreparation> TryPrepareExternalSignature(PdfExternalSignatureOptions? signatureOptions = null, PdfLoadOptions? options = null) =>
        _document.TryPrepareExternalSignature(signatureOptions, options);

    /// <summary>Completes the most recently prepared external signature with encoded CMS bytes.</summary>
    public PdfDocument CompleteExternalSignature(byte[] signatureContents) => _document.CompleteExternalSignature(signatureContents);

    /// <summary>Prepares, delegates signing, and completes an external signature in one operation.</summary>
    public PdfExternalSignatureCompletion SignExternal(IPdfExternalSigner signer, PdfExternalSignatureOptions? options = null) =>
        _document.SignExternal(signer, options);

    /// <summary>
    /// Creates an explicit full-rewrite derivative with invalidated PDF signature fields and revisions removed.
    /// Encrypted sources require owner authorization and the returned derivative is unencrypted.
    /// </summary>
    public PdfUnsignedDerivativeResult CreateUnsignedDerivative(System.Threading.CancellationToken cancellationToken = default) =>
        _document.CreateUnsignedDerivative(cancellationToken);
}
