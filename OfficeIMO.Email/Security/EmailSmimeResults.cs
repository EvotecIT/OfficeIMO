using OfficeIMO.Security;

namespace OfficeIMO.Email;

/// <summary>Result of verifying clear-signed or opaque-signed S/MIME content.</summary>
public sealed class EmailSmimeVerificationResult {
    internal EmailSmimeVerificationResult(
        string providerName,
        EmailProtectionKind protectionKind,
        CmsVerificationResult? cryptography,
        byte[]? signedMimeEntity,
        EmailDocument? signedContent,
        IReadOnlyList<EmailDiagnostic> diagnostics) {
        ProviderName = providerName;
        ProtectionKind = protectionKind;
        Cryptography = cryptography;
        SignedMimeEntity = signedMimeEntity;
        SignedContent = signedContent;
        Diagnostics = diagnostics;
    }

    /// <summary>Explicit security provider that produced the cryptographic evidence.</summary>
    public string ProviderName { get; }
    /// <summary>Protection wrapper detected by the email reader.</summary>
    public EmailProtectionKind ProtectionKind { get; }
    /// <summary>Neutral CMS verification result, or null when no verifiable S/MIME payload was available.</summary>
    public CmsVerificationResult? Cryptography { get; }
    /// <summary>
    /// Exact signed MIME entity bytes extracted from the source. Verification may apply standard CRLF canonicalization
    /// without altering this retained value.
    /// </summary>
    public byte[]? SignedMimeEntity { get; }
    /// <summary>Parsed signed MIME content, when it could be decoded safely.</summary>
    public EmailDocument? SignedContent { get; }
    /// <summary>Email-layer extraction and content-projection diagnostics.</summary>
    public IReadOnlyList<EmailDiagnostic> Diagnostics { get; }
    /// <summary>True when the CMS signature and content digest both validated.</summary>
    public bool IsCryptographicallyValid => Cryptography?.IsCryptographicallyValid == true;
}

/// <summary>Result of decrypting opaque S/MIME EnvelopedData.</summary>
public sealed class EmailSmimeDecryptionResult {
    internal EmailSmimeDecryptionResult(
        string providerName,
        EmailProtectionKind protectionKind,
        CmsDecryptionResult? cryptography,
        byte[]? decryptedMimeEntity,
        EmailDocument? decryptedContent,
        IReadOnlyList<EmailDiagnostic> diagnostics) {
        ProviderName = providerName;
        ProtectionKind = protectionKind;
        Cryptography = cryptography;
        DecryptedMimeEntity = decryptedMimeEntity;
        DecryptedContent = decryptedContent;
        Diagnostics = diagnostics;
    }

    /// <summary>Explicit security provider that produced the cryptographic evidence.</summary>
    public string ProviderName { get; }
    /// <summary>Protection wrapper detected by the email reader.</summary>
    public EmailProtectionKind ProtectionKind { get; }
    /// <summary>Neutral CMS decryption result, or null when no decryptable S/MIME payload was available.</summary>
    public CmsDecryptionResult? Cryptography { get; }
    /// <summary>Exact decrypted MIME entity bytes.</summary>
    public byte[]? DecryptedMimeEntity { get; }
    /// <summary>Parsed decrypted MIME content, when it could be decoded safely.</summary>
    public EmailDocument? DecryptedContent { get; }
    /// <summary>Email-layer extraction and content-projection diagnostics.</summary>
    public IReadOnlyList<EmailDiagnostic> Diagnostics { get; }
    /// <summary>True when CMS decryption succeeded.</summary>
    public bool Decrypted => Cryptography?.Decrypted == true;
}

/// <summary>Ordered stage in an S/MIME protected-entity workflow.</summary>
public enum EmailSmimeProcessingStage {
    /// <summary>CMS EnvelopedData was decrypted first.</summary>
    Decrypt = 0,
    /// <summary>The protected entity exposed by decryption was then signature-verified.</summary>
    Verify = 1
}

/// <summary>Result of decrypting an S/MIME envelope and then verifying protected signed content.</summary>
public sealed class EmailSmimeProcessingResult {
    internal EmailSmimeProcessingResult(string providerName,
        EmailSmimeDecryptionResult decryption, EmailSmimeVerificationResult? verification,
        EmailDocument? content, IReadOnlyList<EmailSmimeProcessingStage> processingOrder,
        IReadOnlyList<EmailDiagnostic> diagnostics) {
        ProviderName = providerName;
        Decryption = decryption;
        Verification = verification;
        Content = content;
        ProcessingOrder = processingOrder;
        Diagnostics = diagnostics;
    }

    /// <summary>Explicit security provider used for every stage.</summary>
    public string ProviderName { get; }
    /// <summary>Outer EnvelopedData decryption evidence.</summary>
    public EmailSmimeDecryptionResult Decryption { get; }
    /// <summary>Inner signature evidence, or null when decrypted content was not signed.</summary>
    public EmailSmimeVerificationResult? Verification { get; }
    /// <summary>Verified signed content when available; otherwise the decrypted MIME content.</summary>
    public EmailDocument? Content { get; }
    /// <summary>Stages actually performed in strict order.</summary>
    public IReadOnlyList<EmailSmimeProcessingStage> ProcessingOrder { get; }
    /// <summary>Combined extraction, trust-policy, and processing diagnostics.</summary>
    public IReadOnlyList<EmailDiagnostic> Diagnostics { get; }
    /// <summary>True when decryption succeeded and any discovered inner signature validated.</summary>
    public bool IsSuccessful => Decryption.Decrypted &&
        (Verification == null || Verification.IsCryptographicallyValid);
}

/// <summary>Stable S/MIME trust-policy diagnostic identifiers.</summary>
public static class EmailSmimeDiagnosticCodes {
    /// <summary>Signer identity evidence was projected.</summary>
    public const string SignerIdentity = "EMAIL_SMIME_SIGNER_IDENTITY";
    /// <summary>Signer certificate-chain outcome.</summary>
    public const string ChainStatus = "EMAIL_SMIME_CHAIN_STATUS";
    /// <summary>Signer revocation outcome.</summary>
    public const string RevocationStatus = "EMAIL_SMIME_REVOCATION_STATUS";
    /// <summary>Signer timestamp outcome.</summary>
    public const string TimestampStatus = "EMAIL_SMIME_TIMESTAMP_STATUS";
    /// <summary>Verification used an offline/no-download trust policy.</summary>
    public const string OfflinePolicy = "EMAIL_SMIME_OFFLINE_POLICY";
    /// <summary>Decryption completed before signature verification.</summary>
    public const string DecryptThenVerify = "EMAIL_SMIME_DECRYPT_THEN_VERIFY";
    /// <summary>The decrypted opaque entity was retained because it was not classified as signed-data.</summary>
    public const string InnerOpaqueNotSigned = "EMAIL_SMIME_INNER_OPAQUE_NOT_SIGNED";
}
