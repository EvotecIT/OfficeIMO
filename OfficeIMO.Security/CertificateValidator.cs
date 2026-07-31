using System.Security.Cryptography.X509Certificates;

namespace OfficeIMO.Security;

/// <summary>Certificate usage profile applied during public trust validation.</summary>
public enum CertificateValidationPurpose {
    /// <summary>A certificate used to sign a document, package, message, or other durable artifact.</summary>
    DocumentSigning,

    /// <summary>A certificate used by an RFC 3161 timestamp authority.</summary>
    TimestampAuthority
}

/// <summary>Platform certificate-chain, revocation, usage, and caller trust-policy result.</summary>
public sealed class CertificateTrustValidationResult {
    internal CertificateTrustValidationResult(
        CertificateValidationResult validation,
        IReadOnlyList<SecurityFinding> findings) {
        Validation = validation;
        Findings = findings;
    }

    /// <summary>Gets the chain and revocation outcome.</summary>
    public CertificateValidationResult Validation { get; }

    /// <summary>Gets stable certificate-profile and trust findings.</summary>
    public IReadOnlyList<SecurityFinding> Findings { get; }
}

/// <summary>Validates an X.509 certificate through the shared OfficeIMO trust-policy engine.</summary>
public static class CertificateValidator {
    /// <summary>
    /// Builds the platform chain, applies revocation and usage policy, and invokes the optional caller trust callback.
    /// </summary>
    /// <param name="certificate">Certificate being validated.</param>
    /// <param name="additionalCertificates">Optional intermediate, root, or peer certificates supplied by the artifact.</param>
    /// <param name="options">Caller-controlled platform trust and revocation policy.</param>
    /// <param name="purpose">Expected certificate usage.</param>
    public static CertificateTrustValidationResult Validate(
        X509Certificate2 certificate,
        IEnumerable<X509Certificate2>? additionalCertificates = null,
        CertificateValidationOptions? options = null,
        CertificateValidationPurpose purpose = CertificateValidationPurpose.DocumentSigning) {
#if NETSTANDARD2_0 || NET472
        if (certificate == null) throw new ArgumentNullException(nameof(certificate));
#else
        ArgumentNullException.ThrowIfNull(certificate);
#endif
        var findings = new List<SecurityFinding>();
        CertificateValidationResult validation = CertificateChainValidator.Validate(
            certificate,
            additionalCertificates ?? Array.Empty<X509Certificate2>(),
            options ?? new CertificateValidationOptions(),
            findings,
            purpose == CertificateValidationPurpose.TimestampAuthority ? "TSA" : "Signer",
            purpose == CertificateValidationPurpose.TimestampAuthority
                ? CertificateUsagePurpose.TimestampAuthority
                : CertificateUsagePurpose.CmsSigner);
        return new CertificateTrustValidationResult(validation, findings.ToArray());
    }
}
