using OfficeIMO.Security;

namespace OfficeIMO.Email;

/// <summary>Result of verifying clear-signed or opaque-signed S/MIME content.</summary>
public sealed class EmailSmimeVerificationResult {
    internal EmailSmimeVerificationResult(
        EmailProtectionKind protectionKind,
        CmsVerificationResult? cryptography,
        byte[]? signedMimeEntity,
        EmailDocument? signedContent,
        IReadOnlyList<EmailDiagnostic> diagnostics) {
        ProtectionKind = protectionKind;
        Cryptography = cryptography;
        SignedMimeEntity = signedMimeEntity;
        SignedContent = signedContent;
        Diagnostics = diagnostics;
    }

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
        EmailProtectionKind protectionKind,
        CmsDecryptionResult? cryptography,
        byte[]? decryptedMimeEntity,
        EmailDocument? decryptedContent,
        IReadOnlyList<EmailDiagnostic> diagnostics) {
        ProtectionKind = protectionKind;
        Cryptography = cryptography;
        DecryptedMimeEntity = decryptedMimeEntity;
        DecryptedContent = decryptedContent;
        Diagnostics = diagnostics;
    }

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
