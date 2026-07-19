using System.IO;
using System.Security.Cryptography.X509Certificates;
using Org.BouncyCastle.Asn1;
using Org.BouncyCastle.Asn1.Cms;
using Org.BouncyCastle.Asn1.Pkcs;
using Org.BouncyCastle.Cms;
using Org.BouncyCastle.Security;
using Org.BouncyCastle.Utilities;
using BcX509Certificate = Org.BouncyCastle.X509.X509Certificate;

namespace OfficeIMO.Security;

/// <summary>Verifies encapsulated and detached CMS SignedData with explicit trust-policy results.</summary>
public static class CmsSignedDataVerifier {
    /// <summary>Verifies an encapsulated CMS SignedData object.</summary>
    public static CmsVerificationResult Verify(byte[] encodedCms, CmsVerificationOptions? options = null) =>
        VerifyCore(encodedCms, null, detachedContentSupplied: false, options);

    /// <summary>Verifies a detached CMS SignedData object against the exact supplied content bytes.</summary>
    public static CmsVerificationResult VerifyDetached(
        byte[] encodedCms,
        byte[] detachedContent,
        CmsVerificationOptions? options = null) {
#if NETSTANDARD2_0 || NET472
        if (detachedContent == null) throw new ArgumentNullException(nameof(detachedContent));
#else
        ArgumentNullException.ThrowIfNull(detachedContent);
#endif
        return VerifyCore(encodedCms, detachedContent, detachedContentSupplied: true, options);
    }

    private static CmsVerificationResult VerifyCore(
        byte[] encodedCms,
        byte[]? detachedContent,
        bool detachedContentSupplied,
        CmsVerificationOptions? options) {
#if NETSTANDARD2_0 || NET472
        if (encodedCms == null) throw new ArgumentNullException(nameof(encodedCms));
#else
        ArgumentNullException.ThrowIfNull(encodedCms);
#endif
        options ??= new CmsVerificationOptions();
        SecurityLimits.EnsureBufferWithinLimit(encodedCms, options.MaxEncodedBytes, nameof(encodedCms));
        if (detachedContent != null) {
            SecurityLimits.EnsureBufferWithinLimit(detachedContent, options.MaxContentBytes, nameof(detachedContent));
        }

        var containerFindings = new List<SecurityFinding>();
        try {
            var decoded = new CmsSignedData(encodedCms);
            bool isDetached = decoded.SignedContent == null;
            byte[]? content;
            CmsSignedData verifiable;
            if (isDetached) {
                content = detachedContent;
                verifiable = detachedContentSupplied
                    ? new CmsSignedData(new CmsProcessableByteArray(detachedContent!), encodedCms)
                    : decoded;
                if (!detachedContentSupplied) {
                    containerFindings.Add(new SecurityFinding(
                        SecurityFindingSeverity.Warning,
                        "DetachedContentMissing",
                        "The CMS object is detached, but no content was supplied for verification."));
                }
            } else {
                try {
                    content = ReadEncapsulatedContent(decoded, options.MaxContentBytes);
                } catch (SecurityContentLimitExceededException exception) {
                    containerFindings.Add(new SecurityFinding(
                        SecurityFindingSeverity.Error,
                        "CmsContentLimitExceeded",
                        exception.Message));
                    return new CmsVerificationResult(
                        parsed: true,
                        isDetached: false,
                        decoded.SignedContentType?.Id,
                        encapsulatedContent: null,
                        Array.Empty<CmsSignerVerificationResult>(),
                        containerFindings);
                }
                verifiable = decoded;
                if (detachedContentSupplied) {
                    containerFindings.Add(new SecurityFinding(
                        SecurityFindingSeverity.Info,
                        "DetachedContentIgnored",
                        "The CMS object contains encapsulated content; the separately supplied content was ignored."));
                }
            }

            List<BcX509Certificate> embedded = verifiable.GetCertificates().EnumerateMatches(null).ToList();
            SecurityLimits.EnsureCountWithinLimit(embedded.Count, options.MaxCertificates, nameof(options.MaxCertificates));
            IList<SignerInformation> signers = verifiable.GetSignerInfos().GetSigners();
            SecurityLimits.EnsureCountWithinLimit(signers.Count, options.MaxSigners, nameof(options.MaxSigners));
            if (signers.Count == 0) {
                containerFindings.Add(new SecurityFinding(
                    SecurityFindingSeverity.Error,
                    "CmsSignerMissing",
                    "The CMS SignedData object contains no signers."));
            }

            var platformEmbedded = CreatePlatformCertificates(embedded, containerFindings);
            try {
                var signerResults = new List<CmsSignerVerificationResult>(signers.Count);
                for (int index = 0; index < signers.Count; index++) {
                    signerResults.Add(VerifySigner(
                        signers[index],
                        index,
                        content,
                        verifiable,
                        embedded,
                        platformEmbedded,
                        options));
                }

                return new CmsVerificationResult(
                    parsed: true,
                    isDetached,
                    verifiable.SignedContentType?.Id,
                    isDetached ? null : content,
                    signerResults,
                    containerFindings);
            } finally {
                foreach (X509Certificate2 certificate in platformEmbedded) certificate.Dispose();
            }
        } catch (Exception exception) when (IsValidationException(exception)) {
            containerFindings.Add(new SecurityFinding(
                SecurityFindingSeverity.Error,
                "CmsMalformed",
                "The CMS SignedData object could not be decoded: " + exception.Message));
            return new CmsVerificationResult(
                parsed: false,
                isDetached: false,
                contentTypeOid: null,
                encapsulatedContent: null,
                Array.Empty<CmsSignerVerificationResult>(),
                containerFindings);
        }
    }

    private static CmsSignerVerificationResult VerifySigner(
        SignerInformation signer,
        int signerIndex,
        byte[]? content,
        CmsSignedData signedData,
        IReadOnlyList<BcX509Certificate> embedded,
        IReadOnlyList<X509Certificate2> platformEmbedded,
        CmsVerificationOptions options) {
        var findings = new List<SecurityFinding>();
        BcX509Certificate? bcSigner = signedData.GetCertificates()
            .EnumerateMatches(signer.SignerID)
            .FirstOrDefault();
        bcSigner ??= FindExtraSignerCertificate(signer, options.CertificateValidation.ExtraCertificates);

        if (bcSigner == null) {
            findings.Add(new SecurityFinding(
                SecurityFindingSeverity.Error,
                "SignerCertificateMissing",
                "No certificate matching the CMS signer identifier was supplied or embedded.",
                signerIndex));
            return CreateMissingCertificateResult(signer, signerIndex, findings);
        }

        using X509Certificate2 platformSigner = PlatformCertificateLoader.Load(bcSigner.GetEncoded());
        SecurityValidationStatus digestStatus = ValidateDigest(signer, content, signerIndex, findings);
        SecurityValidationStatus signatureStatus;
        if (content == null) {
            signatureStatus = SecurityValidationStatus.Indeterminate;
        } else {
            try {
                signatureStatus = signer.Verify(bcSigner)
                    ? SecurityValidationStatus.Valid
                    : SecurityValidationStatus.Invalid;
                if (signatureStatus == SecurityValidationStatus.Invalid) {
                    findings.Add(new SecurityFinding(
                        SecurityFindingSeverity.Error,
                        "CmsSignatureInvalid",
                        "The CMS signature did not verify.",
                        signerIndex));
                }
            } catch (Exception exception) when (IsValidationException(exception)) {
                signatureStatus = SecurityValidationStatus.Invalid;
                findings.Add(new SecurityFinding(
                    SecurityFindingSeverity.Error,
                    "CmsSignatureInvalid",
                    "The CMS signature or signed attributes are invalid: " + exception.Message,
                    signerIndex));
            }
        }

        CertificateValidationResult certificateValidation = CertificateChainValidator.Validate(
            platformSigner,
            platformEmbedded,
            options.CertificateValidation,
            findings,
            "CMS signer",
            signerIndex);
        DateTimeOffset? signingTime = ReadSigningTime(signer.SignedAttributes, signerIndex, findings);
        IReadOnlyList<Rfc3161TimestampVerificationResult> timestamps = options.ValidateTimestamps
            ? VerifyTimestamps(signer, options, signerIndex, findings)
            : Array.Empty<Rfc3161TimestampVerificationResult>();
        SecurityValidationStatus timestampStatus = options.ValidateTimestamps
            ? AggregateTimestampStatus(timestamps)
            : SecurityValidationStatus.NotPerformed;
        DateTimeOffset? timestampTime = timestamps
            .Where(static result => result.Status == SecurityValidationStatus.Valid)
            .Select(static result => result.Timestamp)
            .Where(static value => value.HasValue)
            .OrderByDescending(static value => value)
            .FirstOrDefault();

        return new CmsSignerVerificationResult(
            signerIndex,
            signatureStatus,
            digestStatus,
            certificateValidation,
            timestampStatus,
            bcSigner.GetEncoded(),
            platformSigner.Subject,
            platformSigner.Issuer,
            platformSigner.SerialNumber,
            platformSigner.Thumbprint,
            signer.DigestAlgorithmID.Algorithm.Id,
            signer.SignatureAlgorithm.Algorithm.Id,
            signingTime,
            timestampTime,
            timestamps,
            findings);
    }

    private static SecurityValidationStatus ValidateDigest(
        SignerInformation signer,
        byte[]? content,
        int signerIndex,
        List<SecurityFinding> findings) {
        if (content == null) return SecurityValidationStatus.Indeterminate;
        AttributeTable? signedAttributes = signer.SignedAttributes;
        if (signedAttributes == null) return SecurityValidationStatus.Valid;

        List<Org.BouncyCastle.Asn1.Cms.Attribute> digestAttributes = signedAttributes
            .Where(static attribute => attribute.AttrType.Equals(CmsAttributes.MessageDigest))
            .ToList();
        if (digestAttributes.Count != 1 || digestAttributes[0].AttrValues.Count != 1 ||
            digestAttributes[0].AttrValues[0] is not Asn1OctetString encodedDigest) {
            findings.Add(new SecurityFinding(
                SecurityFindingSeverity.Error,
                digestAttributes.Count == 0 ? "CmsMessageDigestMissing" : "CmsMessageDigestInvalid",
                digestAttributes.Count == 0
                    ? "Signed attributes must contain a message-digest value."
                    : "Signed attributes must contain exactly one well-formed message-digest value.",
                signerIndex));
            return SecurityValidationStatus.Invalid;
        }

        try {
            byte[] calculated = DigestUtilities.CalculateDigest(signer.DigestAlgorithmID.Algorithm, content);
            bool valid = Arrays.FixedTimeEquals(calculated, encodedDigest.GetOctets());
            if (!valid) {
                findings.Add(new SecurityFinding(
                    SecurityFindingSeverity.Error,
                    "CmsContentDigestMismatch",
                    "The signed message-digest does not match the supplied content.",
                    signerIndex));
            }
            return valid ? SecurityValidationStatus.Valid : SecurityValidationStatus.Invalid;
        } catch (Exception exception) when (IsValidationException(exception)) {
            findings.Add(new SecurityFinding(
                SecurityFindingSeverity.Error,
                "CmsDigestUnsupported",
                "The CMS digest algorithm could not be evaluated: " + exception.Message,
                signerIndex));
            return SecurityValidationStatus.Indeterminate;
        }
    }

    private static IReadOnlyList<Rfc3161TimestampVerificationResult> VerifyTimestamps(
        SignerInformation signer,
        CmsVerificationOptions options,
        int signerIndex,
        List<SecurityFinding> findings) {
        AttributeTable? unsignedAttributes = signer.UnsignedAttributes;
        if (unsignedAttributes == null) return Array.Empty<Rfc3161TimestampVerificationResult>();
        var results = new List<Rfc3161TimestampVerificationResult>();
        foreach (Org.BouncyCastle.Asn1.Cms.Attribute attribute in unsignedAttributes) {
            if (!attribute.AttrType.Equals(PkcsObjectIdentifiers.IdAASignatureTimeStampToken)) continue;
            for (int index = 0; index < attribute.AttrValues.Count; index++) {
                byte[] encoded = attribute.AttrValues[index].GetEncoded();
                Rfc3161TimestampVerificationResult result = Rfc3161TimestampVerifier.Verify(
                    encoded,
                    signer.GetSignature(),
                    options.CertificateValidation,
                    Math.Min(options.MaxEncodedBytes, 16L * 1024 * 1024),
                    options.MaxCertificates);
                results.Add(result);
                foreach (SecurityFinding finding in result.Findings) {
                    findings.Add(new SecurityFinding(finding.Severity, finding.Code, finding.Message, signerIndex));
                }
            }
        }
        return results;
    }

    private static DateTimeOffset? ReadSigningTime(
        AttributeTable? signedAttributes,
        int signerIndex,
        List<SecurityFinding> findings) {
        if (signedAttributes == null) return null;
        List<Org.BouncyCastle.Asn1.Cms.Attribute> values = signedAttributes
            .Where(static attribute => attribute.AttrType.Equals(CmsAttributes.SigningTime))
            .ToList();
        if (values.Count == 0) return null;
        if (values.Count != 1 || values[0].AttrValues.Count != 1) {
            findings.Add(new SecurityFinding(
                SecurityFindingSeverity.Warning,
                "CmsSigningTimeInvalid",
                "The signing-time attribute is duplicated or malformed.",
                signerIndex));
            return null;
        }
        try {
            DateTime value = Org.BouncyCastle.Asn1.Cms.Time.GetInstance(values[0].AttrValues[0]).ToDateTime();
            value = value.Kind == DateTimeKind.Utc ? value : DateTime.SpecifyKind(value, DateTimeKind.Utc);
            return new DateTimeOffset(value);
        } catch (Exception exception) when (IsValidationException(exception)) {
            findings.Add(new SecurityFinding(
                SecurityFindingSeverity.Warning,
                "CmsSigningTimeInvalid",
                "The signing-time attribute could not be decoded: " + exception.Message,
                signerIndex));
            return null;
        }
    }

    private static BcX509Certificate? FindExtraSignerCertificate(
        SignerInformation signer,
        X509Certificate2Collection extraCertificates) {
        foreach (X509Certificate2 certificate in extraCertificates) {
            BcX509Certificate candidate = DotNetUtilities.FromX509Certificate(certificate);
            if (signer.SignerID.Match(candidate)) return candidate;
        }
        return null;
    }

    private static List<X509Certificate2> CreatePlatformCertificates(
        List<BcX509Certificate> certificates,
        List<SecurityFinding> findings) {
        var result = new List<X509Certificate2>(certificates.Count);
        try {
            foreach (BcX509Certificate certificate in certificates) {
                result.Add(PlatformCertificateLoader.Load(certificate.GetEncoded()));
            }
            return result;
        } catch (Exception exception) when (IsValidationException(exception)) {
            foreach (X509Certificate2 certificate in result) certificate.Dispose();
            findings.Add(new SecurityFinding(
                SecurityFindingSeverity.Error,
                "CmsCertificateMalformed",
                "An embedded CMS certificate could not be decoded: " + exception.Message));
            throw;
        }
    }

    private static byte[] ReadEncapsulatedContent(CmsSignedData signedData, long maximumBytes) {
        using var stream = new BoundedMemoryStream(maximumBytes);
        signedData.SignedContent.Write(stream);
        return stream.ToArray();
    }

    private static CmsSignerVerificationResult CreateMissingCertificateResult(
        SignerInformation signer,
        int signerIndex,
        IReadOnlyList<SecurityFinding> findings) =>
        new CmsSignerVerificationResult(
            signerIndex,
            SecurityValidationStatus.Indeterminate,
            SecurityValidationStatus.Indeterminate,
            new CertificateValidationResult(
                SecurityValidationStatus.Indeterminate,
                SecurityValidationStatus.NotPerformed,
                Array.Empty<string>()),
            SecurityValidationStatus.NotPerformed,
            null,
            null,
            null,
            null,
            null,
            signer.DigestAlgorithmID.Algorithm.Id,
            signer.SignatureAlgorithm.Algorithm.Id,
            null,
            null,
            Array.Empty<Rfc3161TimestampVerificationResult>(),
            findings);

    private static SecurityValidationStatus AggregateTimestampStatus(
        IReadOnlyList<Rfc3161TimestampVerificationResult> timestamps) {
        if (timestamps.Count == 0) return SecurityValidationStatus.NotPerformed;
        if (timestamps.Any(static result => result.Status == SecurityValidationStatus.Invalid)) {
            return SecurityValidationStatus.Invalid;
        }
        if (timestamps.Any(static result => result.Status == SecurityValidationStatus.Indeterminate)) {
            return SecurityValidationStatus.Indeterminate;
        }
        return SecurityValidationStatus.Valid;
    }

    private static bool IsValidationException(Exception exception) =>
        exception is not OutOfMemoryException &&
        exception is not StackOverflowException &&
        exception is not AccessViolationException;
}
