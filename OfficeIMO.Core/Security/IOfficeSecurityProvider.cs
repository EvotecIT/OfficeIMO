using System.Collections.Generic;
using System.Security.Cryptography.X509Certificates;

namespace OfficeIMO.Security;

/// <summary>
/// Strongly typed cryptographic provider used by OfficeIMO document packages when security features are explicitly enabled.
/// The contract is dependency-free; implementations such as <c>OfficeIMO.Security</c> may use additional cryptographic packages.
/// </summary>
public interface IOfficeSecurityProvider {
    /// <summary>Gets a stable provider name for diagnostics.</summary>
    string Name { get; }

    /// <summary>Creates detached CMS SignedData over the exact supplied content.</summary>
    byte[] SignCmsDetached(
        byte[] content,
        X509Certificate2 signingCertificate,
        CmsSigningOptions? options = null,
        IEnumerable<X509Certificate2>? certificateChain = null);

    /// <summary>Creates encapsulated CMS SignedData containing the supplied content.</summary>
    byte[] SignCmsEncapsulated(
        byte[] content,
        X509Certificate2 signingCertificate,
        CmsSigningOptions? options = null,
        IEnumerable<X509Certificate2>? certificateChain = null);

    /// <summary>Verifies encapsulated CMS SignedData.</summary>
    CmsVerificationResult VerifyCms(
        byte[] encodedCms,
        CmsVerificationOptions? options = null,
        CertificateValidationPurpose signerCertificatePurpose = CertificateValidationPurpose.DocumentSigning);

    /// <summary>Verifies detached CMS SignedData against the exact supplied content.</summary>
    CmsVerificationResult VerifyCmsDetached(
        byte[] encodedCms,
        byte[] detachedContent,
        CmsVerificationOptions? options = null,
        CertificateValidationPurpose signerCertificatePurpose = CertificateValidationPurpose.DocumentSigning);

    /// <summary>
    /// Creates a bounded CMS verification session whose aggregate timestamp limits are shared across every
    /// verification performed by that session.
    /// </summary>
    ICmsVerificationSession CreateCmsVerificationSession(CmsVerificationOptions options);

    /// <summary>Creates CMS EnvelopedData for the supplied recipients.</summary>
    byte[] EncryptCms(
        byte[] content,
        IEnumerable<X509Certificate2> recipients,
        CmsEnvelopeOptions? options = null);

    /// <summary>Decrypts CMS EnvelopedData for a matching recipient certificate.</summary>
    CmsDecryptionResult DecryptCms(
        byte[] encodedCms,
        X509Certificate2 recipientCertificate,
        CmsEnvelopeOptions? options = null);

    /// <summary>Validates certificate trust, revocation, and intended usage under caller policy.</summary>
    CertificateTrustValidationResult ValidateCertificate(
        X509Certificate2 certificate,
        IEnumerable<X509Certificate2>? additionalCertificates = null,
        CertificateValidationOptions? options = null,
        CertificateValidationPurpose purpose = CertificateValidationPurpose.DocumentSigning);

    /// <summary>Verifies an RFC 3161 timestamp token against the exact timestamped bytes.</summary>
    Rfc3161TimestampVerificationResult VerifyTimestamp(
        byte[] encodedToken,
        byte[] timestampedData,
        CertificateValidationOptions? certificateOptions = null,
        long maxEncodedBytes = 16L * 1024L * 1024L,
        int maxCertificates = 64);

    /// <summary>Creates one bounded enveloping XML signature.</summary>
    byte[] CreateXmlSignature(XmlDigitalSignatureCreationRequest request);

    /// <summary>Verifies XML signature math and bounded local-reference processing independently of trust.</summary>
    XmlDigitalSignatureVerificationResult VerifyXmlSignature(XmlDigitalSignatureVerificationRequest request);

    /// <summary>Canonicalizes a bounded XML document using an explicitly supported XML DSig algorithm.</summary>
    byte[] CanonicalizeXml(
        byte[] xml,
        string algorithm,
        string? inclusiveNamespacesPrefixList = null,
        long maxOutputBytes = 16L * 1024L * 1024L);

    /// <summary>Normalizes one bounded DER/BER object, optionally trimming PDF-style trailing zero padding.</summary>
    byte[] NormalizeAsn1Object(
        byte[] encoded,
        bool allowTrailingZeroPadding,
        long maxEncodedBytes);
}

/// <summary>
/// Bounded CMS verification scope used when one document operation verifies several related CMS containers.
/// </summary>
public interface ICmsVerificationSession : System.IDisposable {
    /// <summary>Verifies one encapsulated CMS SignedData container within this operation's shared limits.</summary>
    CmsVerificationResult Verify(
        byte[] encodedCms,
        CertificateValidationPurpose signerCertificatePurpose = CertificateValidationPurpose.DocumentSigning);

    /// <summary>Verifies one detached CMS SignedData container within this operation's shared limits.</summary>
    CmsVerificationResult VerifyDetached(
        byte[] encodedCms,
        byte[] detachedContent,
        CertificateValidationPurpose signerCertificatePurpose = CertificateValidationPurpose.DocumentSigning);
}
