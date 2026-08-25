#if NET8_0_OR_GREATER
using System.Security.Cryptography;
using System.Security.Cryptography.Pkcs;
using System.Security.Cryptography.X509Certificates;

namespace OfficeIMO.Security;

/// <summary>
/// Handles the common detached RSA/SHA CMS shape with the platform parser. More
/// complex structures remain on the Bouncy Castle verifier so this fast path
/// stays small and fail-closed.
/// </summary>
internal static class PlatformCmsSignedDataVerifier {
    private const string CmsDataOid = "1.2.840.113549.1.7.1";
    private const string Sha1Oid = "1.3.14.3.2.26";
    private const string Sha256Oid = "2.16.840.1.101.3.4.2.1";
    private const string Sha384Oid = "2.16.840.1.101.3.4.2.2";
    private const string Sha512Oid = "2.16.840.1.101.3.4.2.3";
    private const string RsaEncryptionOid = "1.2.840.113549.1.1.1";
    private const string Sha1WithRsaOid = "1.2.840.113549.1.1.5";
    private const string Sha256WithRsaOid = "1.2.840.113549.1.1.11";
    private const string Sha384WithRsaOid = "1.2.840.113549.1.1.12";
    private const string Sha512WithRsaOid = "1.2.840.113549.1.1.13";

    internal static bool TryVerifyDetached(
        byte[] encodedCms,
        byte[] detachedContent,
        CmsVerificationOptions options,
        CertificateUsagePurpose signerCertificatePurpose,
        out CmsVerificationResult result) {
        result = null!;
        if (options.ValidateTimestamps || options.CertificateValidation.ValidateChain ||
            options.CertificateValidation.ExtraCertificates.Count != 0 ||
            encodedCms.AsSpan().IndexOf(MessageDigestAttributeOidDer) >= 0) {
            return false;
        }

        try {
            var signedCms = new SignedCms(new ContentInfo(detachedContent), detached: true);
            signedCms.Decode(encodedCms);
            if (!signedCms.Detached ||
                !string.Equals(signedCms.ContentInfo.ContentType.Value, CmsDataOid, StringComparison.Ordinal) ||
                signedCms.SignerInfos.Count != 1 ||
                signedCms.Certificates.Count != 1) {
                return false;
            }

            SecurityLimits.EnsureCountWithinLimit(
                signedCms.SignerInfos.Count,
                options.MaxSigners,
                nameof(options.MaxSigners));
            SecurityLimits.EnsureCountWithinLimit(
                signedCms.Certificates.Count,
                options.MaxCertificates,
                nameof(options.MaxCertificates));

            System.Security.Cryptography.Pkcs.SignerInfo signer = signedCms.SignerInfos[0];
            string digestAlgorithmOid = signer.DigestAlgorithm.Value ?? string.Empty;
            string signatureAlgorithmOid = signer.SignatureAlgorithm.Value ?? string.Empty;
            if (!UsesSupportedRsaSignature(digestAlgorithmOid, signatureAlgorithmOid) ||
                signer.SignedAttributes.Count != 0 ||
                signer.UnsignedAttributes.Count != 0) {
                return false;
            }

            signedCms.CheckSignature(verifySignatureOnly: true);
            X509Certificate2 certificate = signer.Certificate
                ?? throw new CryptographicException("The CMS signer certificate is missing.");
            var findings = new SecurityFindingCollection();
            CertificateValidationResult certificateValidation = CertificateChainValidator.Validate(
                certificate,
                signedCms.Certificates.Cast<X509Certificate2>(),
                options.CertificateValidation,
                ref findings,
                "CMS signer",
                signerCertificatePurpose,
                signerIndex: 0);
            var signerResult = new CmsSignerVerificationResult(
                signerIndex: 0,
                SecurityValidationStatus.Valid,
                SecurityValidationStatus.Valid,
                certificateValidation,
                SecurityValidationStatus.NotPerformed,
                certificate.RawData,
                certificate.Subject,
                certificate.Issuer,
                certificate.SerialNumber,
                certificate.Thumbprint,
                digestAlgorithmOid,
                signatureAlgorithmOid,
                signingTime: null,
                timestampTime: null,
                Array.Empty<Rfc3161TimestampVerificationResult>(),
                findings.Items);
            result = new CmsVerificationResult(
                parsed: true,
                isDetached: true,
                signedCms.ContentInfo.ContentType.Value,
                encapsulatedContent: null,
                authenticodeIndirectData: null,
                new[] { signerResult },
                Array.Empty<SecurityFinding>());
            return true;
        } catch (Exception exception) when (exception is CryptographicException or ArgumentException) {
            return false;
        }
    }

    private static bool UsesSupportedRsaSignature(string digestAlgorithmOid, string signatureAlgorithmOid) {
        if (signatureAlgorithmOid == RsaEncryptionOid) return IsSupportedDigest(digestAlgorithmOid);
        return signatureAlgorithmOid switch {
            Sha1WithRsaOid => digestAlgorithmOid == Sha1Oid,
            Sha256WithRsaOid => digestAlgorithmOid == Sha256Oid,
            Sha384WithRsaOid => digestAlgorithmOid == Sha384Oid,
            Sha512WithRsaOid => digestAlgorithmOid == Sha512Oid,
            _ => false
        };
    }

    private static bool IsSupportedDigest(string digestAlgorithmOid) => digestAlgorithmOid is
        Sha1Oid or Sha256Oid or Sha384Oid or Sha512Oid;

    // DER encoding of 1.2.840.113549.1.9.4. Its presence identifies the
    // required message-digest member of a CMS signed-attributes set. Those
    // richer structures remain on the complete Bouncy Castle verifier.
    private static ReadOnlySpan<byte> MessageDigestAttributeOidDer =>
        [0x06, 0x09, 0x2A, 0x86, 0x48, 0x86, 0xF7, 0x0D, 0x01, 0x09, 0x04];
}
#endif
