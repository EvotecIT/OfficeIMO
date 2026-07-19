using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using Org.BouncyCastle.Asn1;
using Org.BouncyCastle.Asn1.Cms;
using Org.BouncyCastle.Asn1.X509;
using Org.BouncyCastle.Cms;
using Org.BouncyCastle.Security;

namespace OfficeIMO.Security;

/// <summary>Creates interoperable CMS SignedData without exporting an RSA private key.</summary>
public static class CmsSignedDataSigner {
    /// <summary>Creates a detached CMS signature over <paramref name="content"/>.</summary>
    public static byte[] SignDetached(
        byte[] content,
        X509Certificate2 signingCertificate,
        CmsSigningOptions? options = null,
        IEnumerable<X509Certificate2>? certificateChain = null) =>
        Sign(content, signingCertificate, encapsulate: false, options, certificateChain);

    /// <summary>Creates an encapsulated CMS SignedData object containing <paramref name="content"/>.</summary>
    public static byte[] SignEncapsulated(
        byte[] content,
        X509Certificate2 signingCertificate,
        CmsSigningOptions? options = null,
        IEnumerable<X509Certificate2>? certificateChain = null) =>
        Sign(content, signingCertificate, encapsulate: true, options, certificateChain);

    private static byte[] Sign(
        byte[] content,
        X509Certificate2 signingCertificate,
        bool encapsulate,
        CmsSigningOptions? options,
        IEnumerable<X509Certificate2>? certificateChain) {
#if NETSTANDARD2_0 || NET472
        if (content == null) throw new ArgumentNullException(nameof(content));
        if (signingCertificate == null) throw new ArgumentNullException(nameof(signingCertificate));
#else
        ArgumentNullException.ThrowIfNull(content);
        ArgumentNullException.ThrowIfNull(signingCertificate);
#endif
        options ??= new CmsSigningOptions();
        SecurityLimits.EnsureBufferWithinLimit(content, options.MaxContentBytes, nameof(content));

        using RSA? rsa = signingCertificate.GetRSAPrivateKey();
        if (rsa == null) {
            throw new NotSupportedException(
                "CMS signing currently requires an RSA certificate with an accessible private key.");
        }

        Org.BouncyCastle.X509.X509Certificate bcSigner = DotNetUtilities.FromX509Certificate(signingCertificate);
        var signatureFactory = new PlatformRsaSignatureFactory(rsa, options.DigestAlgorithm);
        var attributes = new SignedAttributeGenerator(options);
        var signerBuilder = new SignerInfoGeneratorBuilder()
            .WithSignedAttributeGenerator(attributes);

        var generator = new CmsSignedDataGenerator { UseDefiniteLength = true };
        generator.AddSignerInfoGenerator(signerBuilder.Build(signatureFactory, bcSigner));
        generator.AddCertificate(bcSigner);

        if (options.IncludeCertificateChain && certificateChain != null) {
            var embeddedThumbprints = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
                signingCertificate.Thumbprint ?? string.Empty
            };
            foreach (X509Certificate2 certificate in certificateChain) {
                if (certificate == null || !embeddedThumbprints.Add(certificate.Thumbprint ?? string.Empty)) continue;
                generator.AddCertificate(DotNetUtilities.FromX509Certificate(certificate));
            }
        }

        var processable = new CmsProcessableByteArray(content);
        return generator.Generate(processable, encapsulate).GetEncoded();
    }

    private sealed class SignedAttributeGenerator : CmsAttributeTableGenerator {
        private readonly CmsSigningOptions _options;

        internal SignedAttributeGenerator(CmsSigningOptions options) {
            _options = options;
        }

        public Org.BouncyCastle.Asn1.Cms.AttributeTable GetAttributes(
            IDictionary<CmsAttributeTableParameter, object> parameters) {
            var attributes = new Asn1EncodableVector();
            if (parameters.TryGetValue(CmsAttributeTableParameter.ContentType, out object? contentType) &&
                contentType is DerObjectIdentifier contentTypeOid) {
                attributes.Add(new Org.BouncyCastle.Asn1.Cms.Attribute(
                    CmsAttributes.ContentType,
                    new DerSet(contentTypeOid)));
            }

            if (!parameters.TryGetValue(CmsAttributeTableParameter.Digest, out object? digestValue) ||
                digestValue is not byte[] digest) {
                throw new InvalidOperationException("The CMS generator did not provide a content digest.");
            }
            attributes.Add(new Org.BouncyCastle.Asn1.Cms.Attribute(
                CmsAttributes.MessageDigest,
                new DerSet(new DerOctetString(digest))));

            if (_options.IncludeSigningTime) {
                DateTimeOffset signingTime = _options.SigningTime ?? DateTimeOffset.UtcNow;
                attributes.Add(new Org.BouncyCastle.Asn1.Cms.Attribute(
                    CmsAttributes.SigningTime,
                    new DerSet(new Org.BouncyCastle.Asn1.Cms.Time(signingTime.UtcDateTime))));
            }

            if (parameters.TryGetValue(CmsAttributeTableParameter.DigestAlgorithmIdentifier, out object? digestAlgorithm) &&
                digestAlgorithm is AlgorithmIdentifier digestAlgorithmIdentifier &&
                parameters.TryGetValue(CmsAttributeTableParameter.SignatureAlgorithmIdentifier, out object? signatureAlgorithm) &&
                signatureAlgorithm is AlgorithmIdentifier signatureAlgorithmIdentifier) {
                attributes.Add(new Org.BouncyCastle.Asn1.Cms.Attribute(
                    CmsAttributes.CmsAlgorithmProtect,
                    new DerSet(new CmsAlgorithmProtection(
                        digestAlgorithmIdentifier,
                        CmsAlgorithmProtection.Signature,
                        signatureAlgorithmIdentifier))));
            }

            return new Org.BouncyCastle.Asn1.Cms.AttributeTable(attributes);
        }
    }
}
