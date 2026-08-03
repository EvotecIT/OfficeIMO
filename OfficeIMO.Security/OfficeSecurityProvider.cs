using System.Collections.Generic;
using System.Security.Cryptography.X509Certificates;

namespace OfficeIMO.Security;

/// <summary>Default OfficeIMO cryptographic provider backed by the dependencies carried by this package.</summary>
public sealed class OfficeSecurityProvider : IOfficeSecurityProvider {
    /// <summary>Gets a reusable stateless provider instance.</summary>
    public static OfficeSecurityProvider Default { get; } = new OfficeSecurityProvider();

    /// <inheritdoc />
    public string Name => "OfficeIMO.Security";

    /// <inheritdoc />
    public byte[] SignCmsDetached(
        byte[] content,
        X509Certificate2 signingCertificate,
        CmsSigningOptions? options = null,
        IEnumerable<X509Certificate2>? certificateChain = null) =>
        CmsSignedDataSigner.SignDetached(content, signingCertificate, options, certificateChain);

    /// <inheritdoc />
    public byte[] SignCmsEncapsulated(
        byte[] content,
        X509Certificate2 signingCertificate,
        CmsSigningOptions? options = null,
        IEnumerable<X509Certificate2>? certificateChain = null) =>
        CmsSignedDataSigner.SignEncapsulated(content, signingCertificate, options, certificateChain);

    /// <inheritdoc />
    public CmsVerificationResult VerifyCms(
        byte[] encodedCms,
        CmsVerificationOptions? options = null,
        CertificateValidationPurpose signerCertificatePurpose = CertificateValidationPurpose.DocumentSigning) =>
        CmsSignedDataVerifier.Verify(encodedCms, options ?? new CmsVerificationOptions(), signerCertificatePurpose);

    /// <inheritdoc />
    public CmsVerificationResult VerifyCmsDetached(
        byte[] encodedCms,
        byte[] detachedContent,
        CmsVerificationOptions? options = null,
        CertificateValidationPurpose signerCertificatePurpose = CertificateValidationPurpose.DocumentSigning) =>
        CmsSignedDataVerifier.VerifyDetached(
            encodedCms,
            detachedContent,
            options ?? new CmsVerificationOptions(),
            signerCertificatePurpose);

    /// <inheritdoc />
    public ICmsVerificationSession CreateCmsVerificationSession(CmsVerificationOptions options) {
#if NETSTANDARD2_0 || NET472
        if (options == null) throw new ArgumentNullException(nameof(options));
#else
        ArgumentNullException.ThrowIfNull(options);
#endif
        CmsSignedDataVerifier.ValidateOptions(options);
        return new CmsVerificationSession(options);
    }

    /// <inheritdoc />
    public byte[] EncryptCms(
        byte[] content,
        IEnumerable<X509Certificate2> recipients,
        CmsEnvelopeOptions? options = null) =>
        CmsEnvelopedDataService.Encrypt(content, recipients, options);

    /// <inheritdoc />
    public CmsDecryptionResult DecryptCms(
        byte[] encodedCms,
        X509Certificate2 recipientCertificate,
        CmsEnvelopeOptions? options = null) =>
        CmsEnvelopedDataService.Decrypt(encodedCms, recipientCertificate, options);

    /// <inheritdoc />
    public CertificateTrustValidationResult ValidateCertificate(
        X509Certificate2 certificate,
        IEnumerable<X509Certificate2>? additionalCertificates = null,
        CertificateValidationOptions? options = null,
        CertificateValidationPurpose purpose = CertificateValidationPurpose.DocumentSigning) =>
        CertificateValidator.Validate(certificate, additionalCertificates, options, purpose);

    /// <inheritdoc />
    public Rfc3161TimestampVerificationResult VerifyTimestamp(
        byte[] encodedToken,
        byte[] timestampedData,
        CertificateValidationOptions? certificateOptions = null,
        long maxEncodedBytes = 16L * 1024L * 1024L,
        int maxCertificates = 64) =>
        Rfc3161TimestampVerifier.Verify(
            encodedToken,
            timestampedData,
            certificateOptions,
            maxEncodedBytes,
            maxCertificates);

    /// <inheritdoc />
    public byte[] CreateXmlSignature(XmlDigitalSignatureCreationRequest request) =>
        XmlDigitalSignatureService.Create(request);

    /// <inheritdoc />
    public XmlDigitalSignatureVerificationResult VerifyXmlSignature(XmlDigitalSignatureVerificationRequest request) =>
        XmlDigitalSignatureService.Verify(request);

    /// <inheritdoc />
    public byte[] CanonicalizeXml(
        byte[] xml,
        string algorithm,
        string? inclusiveNamespacesPrefixList = null,
        long maxOutputBytes = 16L * 1024L * 1024L) =>
        XmlDigitalSignatureService.Canonicalize(
            xml,
            algorithm,
            inclusiveNamespacesPrefixList,
            maxOutputBytes);

    /// <inheritdoc />
    public byte[] NormalizeAsn1Object(
        byte[] encoded,
        bool allowTrailingZeroPadding,
        long maxEncodedBytes) =>
        SecurityEncoding.NormalizeSingleAsn1Object(encoded, allowTrailingZeroPadding, maxEncodedBytes);

    private sealed class CmsVerificationSession : ICmsVerificationSession {
        private readonly CmsVerificationOptions _options;
        private readonly CmsSignedDataVerifier.TimestampVerificationBudget _timestampBudget;
        private bool _disposed;

        internal CmsVerificationSession(CmsVerificationOptions options) {
            _options = options;
            _timestampBudget = new CmsSignedDataVerifier.TimestampVerificationBudget(options);
        }

        public CmsVerificationResult Verify(
            byte[] encodedCms,
            CertificateValidationPurpose signerCertificatePurpose = CertificateValidationPurpose.DocumentSigning) {
            ThrowIfDisposed();
            return CmsSignedDataVerifier.Verify(
                encodedCms,
                _options,
                _timestampBudget,
                signerCertificatePurpose);
        }

        public CmsVerificationResult VerifyDetached(
            byte[] encodedCms,
            byte[] detachedContent,
            CertificateValidationPurpose signerCertificatePurpose = CertificateValidationPurpose.DocumentSigning) {
            ThrowIfDisposed();
            return CmsSignedDataVerifier.VerifyDetached(
                encodedCms,
                detachedContent,
                _options,
                _timestampBudget,
                signerCertificatePurpose);
        }

        public void Dispose() => _disposed = true;

        private void ThrowIfDisposed() {
#if NETSTANDARD2_0 || NET472
            if (_disposed) throw new ObjectDisposedException(nameof(CmsVerificationSession));
#else
            ObjectDisposedException.ThrowIf(_disposed, this);
#endif
        }
    }
}
