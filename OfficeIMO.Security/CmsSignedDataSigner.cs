using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using Org.BouncyCastle.Asn1;
using Org.BouncyCastle.Asn1.Cms;
using Org.BouncyCastle.Asn1.X509;
using Org.BouncyCastle.Cms;
using Org.BouncyCastle.Security;
using Org.BouncyCastle.Operators.Utilities;

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
        string contentTypeOid = options.ContentTypeOid ?? CmsObjectIdentifiers.Data.Id;
        try {
            _ = new DerObjectIdentifier(contentTypeOid);
        } catch (Exception exception) when (exception is ArgumentException or FormatException) {
            throw new ArgumentException(
                "The CMS content type must be a valid ASN.1 object identifier.",
                nameof(options),
                exception);
        }

        using RSA? rsa = signingCertificate.GetRSAPrivateKey();
        if (rsa == null) {
            throw new NotSupportedException(
                "CMS signing currently requires an RSA certificate with an accessible private key.");
        }

        Org.BouncyCastle.X509.X509Certificate bcSigner = DotNetUtilities.FromX509Certificate(signingCertificate);
        AlgorithmIdentifier signatureAlgorithm = DefaultSignatureAlgorithmFinder.Instance.Find(
            GetSignatureAlgorithmName(options.DigestAlgorithm));
        AlgorithmIdentifier digestAlgorithm = DefaultDigestAlgorithmFinder.Instance.Find(signatureAlgorithm);
        byte[] contentDigest = ComputeDigest(content, options.DigestAlgorithm);
        var attributeParameters = new Dictionary<CmsAttributeTableParameter, object> {
            [CmsAttributeTableParameter.ContentType] = new DerObjectIdentifier(contentTypeOid),
            [CmsAttributeTableParameter.Digest] = contentDigest,
            [CmsAttributeTableParameter.DigestAlgorithmIdentifier] = digestAlgorithm,
            [CmsAttributeTableParameter.SignatureAlgorithmIdentifier] = signatureAlgorithm
        };
        Org.BouncyCastle.Asn1.Cms.AttributeTable attributeTable =
            new SignedAttributeGenerator(options).GetAttributes(attributeParameters);
        var signedAttributes = new DerSet(attributeTable.ToAsn1EncodableVector());
        byte[] signedAttributeBytes = signedAttributes.GetEncoded(Asn1Encodable.Der);
        byte[] signature = rsa.SignData(
            signedAttributeBytes,
            options.DigestAlgorithm,
            RSASignaturePadding.Pkcs1);

        var signerIdentifier = new SignerIdentifier(
            new IssuerAndSerialNumber(bcSigner.IssuerDN, bcSigner.SerialNumber));
        var signerInfo = new SignerInfo(
            signerIdentifier,
            digestAlgorithm,
            signedAttributes,
            signatureAlgorithm,
            new DerOctetString(signature),
            unauthenticatedAttributes: null);

        var certificates = new List<Asn1Encodable> { bcSigner.CertificateStructure };
        if (options.IncludeCertificateChain && certificateChain != null) {
            var embeddedThumbprints = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
                signingCertificate.Thumbprint ?? string.Empty
            };
            foreach (X509Certificate2 certificate in certificateChain) {
                if (certificate == null || !embeddedThumbprints.Add(certificate.Thumbprint ?? string.Empty)) continue;
                certificates.Add(DotNetUtilities.FromX509Certificate(certificate).CertificateStructure);
            }
        }

        var encapsulatedContent = new ContentInfo(
            new DerObjectIdentifier(contentTypeOid),
            encapsulate ? new DerOctetString(content) : null);
        var signedData = new SignedData(
            new DerSet(digestAlgorithm),
            encapsulatedContent,
            new DerSet(certificates),
            crls: null,
            new DerSet(signerInfo));
        return new ContentInfo(CmsObjectIdentifiers.SignedData, signedData)
            .GetEncoded(Asn1Encodable.Der);
    }

    private static string GetSignatureAlgorithmName(HashAlgorithmName digestAlgorithm) {
        if (digestAlgorithm == HashAlgorithmName.SHA256) return "SHA256WITHRSA";
        if (digestAlgorithm == HashAlgorithmName.SHA384) return "SHA384WITHRSA";
        if (digestAlgorithm == HashAlgorithmName.SHA512) return "SHA512WITHRSA";
        if (digestAlgorithm == HashAlgorithmName.SHA1) return "SHA1WITHRSA";
        throw new NotSupportedException("CMS RSA signing supports SHA-1, SHA-256, SHA-384, and SHA-512.");
    }

    private static byte[] ComputeDigest(byte[] content, HashAlgorithmName digestAlgorithm) {
        using HashAlgorithm algorithm = digestAlgorithm == HashAlgorithmName.SHA256
            ? SHA256.Create()
            : digestAlgorithm == HashAlgorithmName.SHA384
                ? SHA384.Create()
                : digestAlgorithm == HashAlgorithmName.SHA512
                    ? SHA512.Create()
                    : digestAlgorithm == HashAlgorithmName.SHA1
                        ? CreateSha1DigestForLegacyCmsCompatibility()
                        : throw new NotSupportedException(
                            "CMS RSA signing supports SHA-1, SHA-256, SHA-384, and SHA-512.");
        return algorithm.ComputeHash(content);
    }

#pragma warning disable CA5350 // Caller-selected SHA-1 is retained only for legacy CMS interoperability.
    private static SHA1 CreateSha1DigestForLegacyCmsCompatibility() => SHA1.Create();
#pragma warning restore CA5350

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
