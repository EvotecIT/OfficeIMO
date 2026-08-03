using System.Security.Cryptography.X509Certificates;

namespace OfficeIMO.Security;

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
        CertificateUsagePurpose usagePurpose = purpose switch {
            CertificateValidationPurpose.TimestampAuthority => CertificateUsagePurpose.TimestampAuthority,
            CertificateValidationPurpose.EmailSigning => CertificateUsagePurpose.CmsSigner,
            _ => CertificateUsagePurpose.DocumentSigner
        };
        CertificateValidationResult validation = CertificateChainValidator.Validate(
            certificate,
            additionalCertificates ?? Array.Empty<X509Certificate2>(),
            options ?? new CertificateValidationOptions(),
            findings,
            purpose == CertificateValidationPurpose.TimestampAuthority ? "TSA" : "Signer",
            usagePurpose);
        return new CertificateTrustValidationResult(validation, findings.ToArray());
    }
}
