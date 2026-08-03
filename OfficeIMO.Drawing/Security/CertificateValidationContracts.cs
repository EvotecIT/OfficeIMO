using System.Collections.Generic;

namespace OfficeIMO.Security;

/// <summary>Certificate usage profile applied during public trust validation.</summary>
public enum CertificateValidationPurpose {
    /// <summary>A certificate used to sign a document, package, message, or other durable artifact.</summary>
    DocumentSigning = 0,

    /// <summary>A certificate used by an RFC 3161 timestamp authority.</summary>
    TimestampAuthority = 1,

    /// <summary>A certificate used to sign S/MIME email and therefore permitted to declare the email-protection EKU.</summary>
    EmailSigning = 2
}

/// <summary>Platform certificate-chain, revocation, usage, and caller trust-policy result.</summary>
public sealed class CertificateTrustValidationResult {
    /// <summary>Creates certificate trust evidence for a provider implementation.</summary>
    public CertificateTrustValidationResult(
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
