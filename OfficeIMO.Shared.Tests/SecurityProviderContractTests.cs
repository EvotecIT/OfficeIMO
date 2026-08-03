using System.Security.Cryptography.X509Certificates;
using OfficeIMO.Security;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class SecurityProviderContractTests {
    [Fact]
    public void CertificateValidationPurposeKeepsStableNumericValues() {
        Assert.Equal(0, (int)CertificateValidationPurpose.DocumentSigning);
        Assert.Equal(1, (int)CertificateValidationPurpose.TimestampAuthority);
        Assert.Equal(2, (int)CertificateValidationPurpose.EmailSigning);
    }

    [Fact]
    public void DependencyFreeProviderContractCanBeImplementedByAnotherAssembly() {
        IOfficeSecurityProvider provider = new ExternalProviderStub();

        Assert.Equal("External test provider", provider.Name);
        Assert.False(provider.VerifyCms(Array.Empty<byte>()).Parsed);
        Assert.False(provider.DecryptCms(Array.Empty<byte>(), null!).Decrypted);
        Assert.Equal(
            SecurityValidationStatus.NotPerformed,
            provider.ValidateCertificate(null!).Validation.ChainStatus);
        Assert.Equal(
            SecurityValidationStatus.NotPerformed,
            provider.VerifyXmlSignature(new XmlDigitalSignatureVerificationRequest(
                Array.Empty<byte>(),
                Array.Empty<X509Certificate2>())).Status);
    }

    private sealed class ExternalProviderStub : IOfficeSecurityProvider {
        private static readonly CertificateValidationResult EmptyCertificateValidation = new(
            SecurityValidationStatus.NotPerformed,
            SecurityValidationStatus.NotPerformed,
            Array.Empty<string>());

        public string Name => "External test provider";

        public byte[] SignCmsDetached(byte[] content, X509Certificate2 signingCertificate,
            CmsSigningOptions? options = null, IEnumerable<X509Certificate2>? certificateChain = null) =>
            Array.Empty<byte>();

        public byte[] SignCmsEncapsulated(byte[] content, X509Certificate2 signingCertificate,
            CmsSigningOptions? options = null, IEnumerable<X509Certificate2>? certificateChain = null) =>
            Array.Empty<byte>();

        public CmsVerificationResult VerifyCms(byte[] encodedCms, CmsVerificationOptions? options = null,
            CertificateValidationPurpose signerCertificatePurpose = CertificateValidationPurpose.DocumentSigning) =>
            EmptyCmsResult(isDetached: false);

        public CmsVerificationResult VerifyCmsDetached(byte[] encodedCms, byte[] detachedContent,
            CmsVerificationOptions? options = null,
            CertificateValidationPurpose signerCertificatePurpose = CertificateValidationPurpose.DocumentSigning) =>
            EmptyCmsResult(isDetached: true);

        public ICmsVerificationSession CreateCmsVerificationSession(CmsVerificationOptions options) =>
            new ExternalSession();

        public byte[] EncryptCms(byte[] content, IEnumerable<X509Certificate2> recipients,
            CmsEnvelopeOptions? options = null) => Array.Empty<byte>();

        public CmsDecryptionResult DecryptCms(byte[] encodedCms, X509Certificate2 recipientCertificate,
            CmsEnvelopeOptions? options = null) =>
            new(false, false, null, null, null, Array.Empty<SecurityFinding>());

        public CertificateTrustValidationResult ValidateCertificate(X509Certificate2 certificate,
            IEnumerable<X509Certificate2>? additionalCertificates = null,
            CertificateValidationOptions? options = null,
            CertificateValidationPurpose purpose = CertificateValidationPurpose.DocumentSigning) =>
            new(EmptyCertificateValidation, Array.Empty<SecurityFinding>());

        public Rfc3161TimestampVerificationResult VerifyTimestamp(byte[] encodedToken, byte[] timestampedData,
            CertificateValidationOptions? certificateOptions = null, long maxEncodedBytes = 16L * 1024L * 1024L,
            int maxCertificates = 64) =>
            new(SecurityValidationStatus.NotPerformed, null, null, null, null,
                EmptyCertificateValidation, Array.Empty<SecurityFinding>());

        public byte[] CreateXmlSignature(XmlDigitalSignatureCreationRequest request) => Array.Empty<byte>();

        public XmlDigitalSignatureVerificationResult VerifyXmlSignature(
            XmlDigitalSignatureVerificationRequest request) =>
            new(SecurityValidationStatus.NotPerformed, Array.Empty<X509Certificate2>(), Array.Empty<SecurityFinding>());

        public byte[] CanonicalizeXml(byte[] xml, string algorithm,
            string? inclusiveNamespacesPrefixList = null, long maxOutputBytes = 16L * 1024L * 1024L) =>
            (byte[])xml.Clone();

        public byte[] NormalizeAsn1Object(byte[] encoded, bool allowTrailingZeroPadding,
            long maxEncodedBytes) => (byte[])encoded.Clone();

        private static CmsVerificationResult EmptyCmsResult(bool isDetached) =>
            new(false, isDetached, null, null, null,
                Array.Empty<CmsSignerVerificationResult>(), Array.Empty<SecurityFinding>());

        private sealed class ExternalSession : ICmsVerificationSession {
            public CmsVerificationResult Verify(byte[] encodedCms,
                CertificateValidationPurpose signerCertificatePurpose = CertificateValidationPurpose.DocumentSigning) =>
                EmptyCmsResult(isDetached: false);

            public CmsVerificationResult VerifyDetached(byte[] encodedCms, byte[] detachedContent,
                CertificateValidationPurpose signerCertificatePurpose = CertificateValidationPurpose.DocumentSigning) =>
                EmptyCmsResult(isDetached: true);

            public void Dispose() { }
        }
    }
}
