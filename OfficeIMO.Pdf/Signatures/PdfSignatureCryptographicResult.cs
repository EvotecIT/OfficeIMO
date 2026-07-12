namespace OfficeIMO.Pdf;

/// <summary>Cryptographic result for one PDF signature, separated by validation dimension.</summary>
public sealed class PdfSignatureCryptographicResult {
    /// <summary>Creates a provider result.</summary>
    public PdfSignatureCryptographicResult(
        string providerName,
        PdfCryptographicValidationStatus mathematicalSignatureStatus,
        PdfCryptographicValidationStatus messageDigestStatus,
        PdfCryptographicValidationStatus certificateChainStatus,
        PdfCryptographicValidationStatus revocationStatus,
        PdfCryptographicValidationStatus timestampStatus,
        string? signerSubject = null,
        string? signerIssuer = null,
        string? signerSerialNumber = null,
        string? signerThumbprint = null,
        DateTimeOffset? signingTime = null,
        DateTimeOffset? timestampTime = null,
        IReadOnlyList<PdfSignatureCryptographicFinding>? findings = null) {
        if (string.IsNullOrWhiteSpace(providerName)) throw new ArgumentException("Provider name cannot be empty.", nameof(providerName));
        ProviderName = providerName;
        MathematicalSignatureStatus = mathematicalSignatureStatus;
        MessageDigestStatus = messageDigestStatus;
        CertificateChainStatus = certificateChainStatus;
        RevocationStatus = revocationStatus;
        TimestampStatus = timestampStatus;
        SignerSubject = signerSubject;
        SignerIssuer = signerIssuer;
        SignerSerialNumber = signerSerialNumber;
        SignerThumbprint = signerThumbprint;
        SigningTime = signingTime;
        TimestampTime = timestampTime;
        Findings = findings ?? Array.Empty<PdfSignatureCryptographicFinding>();
    }

    /// <summary>Provider that produced this result.</summary>
    public string ProviderName { get; }

    /// <summary>Public-key signature math status.</summary>
    public PdfCryptographicValidationStatus MathematicalSignatureStatus { get; }

    /// <summary>Signed-content digest or timestamp message-imprint status.</summary>
    public PdfCryptographicValidationStatus MessageDigestStatus { get; }

    /// <summary>Signer or timestamp-authority certificate-chain status.</summary>
    public PdfCryptographicValidationStatus CertificateChainStatus { get; }

    /// <summary>OCSP/CRL policy outcome.</summary>
    public PdfCryptographicValidationStatus RevocationStatus { get; }

    /// <summary>RFC 3161 or signature-timestamp status.</summary>
    public PdfCryptographicValidationStatus TimestampStatus { get; }

    /// <summary>Signer certificate subject, when available.</summary>
    public string? SignerSubject { get; }

    /// <summary>Signer certificate issuer, when available.</summary>
    public string? SignerIssuer { get; }

    /// <summary>Signer certificate serial number, when available.</summary>
    public string? SignerSerialNumber { get; }

    /// <summary>Signer certificate thumbprint, when available.</summary>
    public string? SignerThumbprint { get; }

    /// <summary>CMS signing time, when present and readable.</summary>
    public DateTimeOffset? SigningTime { get; }

    /// <summary>Validated RFC 3161 timestamp time, when available.</summary>
    public DateTimeOffset? TimestampTime { get; }

    /// <summary>Provider findings mapped into the aggregate PDF validation report.</summary>
    public IReadOnlyList<PdfSignatureCryptographicFinding> Findings { get; }

    /// <summary>True when signature math and signed-content digest both validated.</summary>
    public bool IsMathematicallyValid =>
        MathematicalSignatureStatus == PdfCryptographicValidationStatus.Valid &&
        MessageDigestStatus == PdfCryptographicValidationStatus.Valid;
}
