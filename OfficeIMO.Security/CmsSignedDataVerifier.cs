using System.IO;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using Org.BouncyCastle.Asn1;
using Org.BouncyCastle.Asn1.Cms;
using Org.BouncyCastle.Asn1.Pkcs;
using Org.BouncyCastle.Cms;
using Org.BouncyCastle.Security;
using Org.BouncyCastle.Utilities;
using Org.BouncyCastle.Utilities.Collections;
using BcX509Certificate = Org.BouncyCastle.X509.X509Certificate;

namespace OfficeIMO.Security;

/// <summary>Verifies encapsulated and detached CMS SignedData with explicit trust-policy results.</summary>
public static class CmsSignedDataVerifier {
    /// <summary>Verifies an encapsulated CMS SignedData object.</summary>
    public static CmsVerificationResult Verify(byte[] encodedCms, CmsVerificationOptions? options = null) =>
        VerifyCore(encodedCms, null, detachedContentSupplied: false, options, timestampBudget: null,
            CertificateUsagePurpose.CmsSigner);

    internal static CmsVerificationResult Verify(
        byte[] encodedCms,
        CmsVerificationOptions options,
        CertificateValidationPurpose signerCertificatePurpose) =>
        VerifyCore(encodedCms, null, detachedContentSupplied: false, options, timestampBudget: null,
            MapCertificatePurpose(signerCertificatePurpose));

    internal static CmsVerificationResult Verify(
        byte[] encodedCms,
        CmsVerificationOptions options,
        TimestampVerificationBudget timestampBudget) =>
        VerifyCore(encodedCms, null, detachedContentSupplied: false, options, timestampBudget,
            CertificateUsagePurpose.CmsSigner);

    internal static CmsVerificationResult Verify(
        byte[] encodedCms,
        CmsVerificationOptions options,
        TimestampVerificationBudget timestampBudget,
        CertificateValidationPurpose signerCertificatePurpose) =>
        VerifyCore(encodedCms, null, detachedContentSupplied: false, options, timestampBudget,
            MapCertificatePurpose(signerCertificatePurpose));

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
        return VerifyCore(encodedCms, detachedContent, detachedContentSupplied: true, options, timestampBudget: null,
            CertificateUsagePurpose.CmsSigner);
    }

    internal static CmsVerificationResult VerifyDetached(
        byte[] encodedCms,
        byte[] detachedContent,
        CmsVerificationOptions options,
        CertificateValidationPurpose signerCertificatePurpose) {
#if NETSTANDARD2_0 || NET472
        if (detachedContent == null) throw new ArgumentNullException(nameof(detachedContent));
#else
        ArgumentNullException.ThrowIfNull(detachedContent);
#endif
        return VerifyCore(encodedCms, detachedContent, detachedContentSupplied: true, options, timestampBudget: null,
            MapCertificatePurpose(signerCertificatePurpose));
    }

    internal static CmsVerificationResult VerifyDetached(
        byte[] encodedCms,
        byte[] detachedContent,
        CmsVerificationOptions options,
        TimestampVerificationBudget timestampBudget,
        CertificateValidationPurpose signerCertificatePurpose) {
#if NETSTANDARD2_0 || NET472
        if (detachedContent == null) throw new ArgumentNullException(nameof(detachedContent));
#else
        ArgumentNullException.ThrowIfNull(detachedContent);
#endif
        return VerifyCore(encodedCms, detachedContent, detachedContentSupplied: true, options, timestampBudget,
            MapCertificatePurpose(signerCertificatePurpose));
    }

    private static CmsVerificationResult VerifyCore(
        byte[] encodedCms,
        byte[]? detachedContent,
        bool detachedContentSupplied,
        CmsVerificationOptions? options,
        TimestampVerificationBudget? timestampBudget,
        CertificateUsagePurpose signerCertificatePurpose) {
#if NETSTANDARD2_0 || NET472
        if (encodedCms == null) throw new ArgumentNullException(nameof(encodedCms));
#else
        ArgumentNullException.ThrowIfNull(encodedCms);
#endif
        options ??= new CmsVerificationOptions();
        ValidateOptions(options);
        SecurityLimits.EnsureBufferWithinLimit(encodedCms, options.MaxEncodedBytes, nameof(encodedCms));
        if (detachedContent != null) {
            SecurityLimits.EnsureBufferWithinLimit(detachedContent, options.MaxContentBytes, nameof(detachedContent));
        }

        var containerFindings = new List<SecurityFinding>();
        try {
            var decoded = new CmsSignedData(encodedCms);
            bool isDetached = decoded.SignedContent == null;
            byte[]? content;
            if (isDetached) {
                content = detachedContent;
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
                        authenticodeIndirectData: null,
                        Array.Empty<CmsSignerVerificationResult>(),
                        containerFindings);
                }
                if (detachedContentSupplied) {
                    containerFindings.Add(new SecurityFinding(
                        SecurityFindingSeverity.Info,
                        "DetachedContentIgnored",
                        "The CMS object contains encapsulated content; the separately supplied content was ignored."));
                }
            }

            List<BcX509Certificate> embedded = decoded.GetCertificates().EnumerateMatches(null).ToList();
            SecurityLimits.EnsureCountWithinLimit(embedded.Count, options.MaxCertificates, nameof(options.MaxCertificates));
            IList<SignerInformation> signers = decoded.GetSignerInfos().GetSigners();
            bool allSignersUseRsaFastPath = signers.All(static signer => CanUseRsaFastPath(signer));
            if (isDetached && detachedContentSupplied && !allSignersUseRsaFastPath) {
                signers = new CmsSignedData(new CmsProcessableByteArray(detachedContent!), encodedCms)
                    .GetSignerInfos()
                    .GetSigners();
            }
            SecurityLimits.EnsureCountWithinLimit(signers.Count, options.MaxSigners, nameof(options.MaxSigners));
            if (signers.Count == 0) {
                containerFindings.Add(new SecurityFinding(
                    SecurityFindingSeverity.Error,
                    "CmsSignerMissing",
                    "The CMS SignedData object contains no signers."));
            }

            var platformEmbedded = CreatePlatformCertificates(embedded, containerFindings, out List<byte[]> embeddedEncodings);
            try {
                var signerResults = new List<CmsSignerVerificationResult>(signers.Count);
                TimestampVerificationBudget? effectiveTimestampBudget = options.ValidateTimestamps
                    ? timestampBudget ?? new TimestampVerificationBudget(options)
                    : null;
                for (int index = 0; index < signers.Count; index++) {
                    signerResults.Add(VerifySigner(
                        signers[index],
                        index,
                        content,
                        allSignersUseRsaFastPath || CanUseRsaFastPath(signers[index]),
                        embedded,
                        embeddedEncodings,
                        platformEmbedded,
                        options,
                        effectiveTimestampBudget,
                        signerCertificatePurpose));
                }

                AuthenticodeIndirectDataInfo? authenticode = TryReadAuthenticodeIndirectData(
                    decoded.SignedContentType?.Id, content, containerFindings);
                return new CmsVerificationResult(
                    parsed: true,
                    isDetached,
                    decoded.SignedContentType?.Id,
                    isDetached ? null : content,
                    authenticode,
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
                authenticodeIndirectData: null,
                Array.Empty<CmsSignerVerificationResult>(),
                containerFindings);
        }
    }

    private static CmsSignerVerificationResult VerifySigner(
        SignerInformation signer,
        int signerIndex,
        byte[]? content,
        bool useRsaFastPath,
        List<BcX509Certificate> embedded,
        List<byte[]> embeddedEncodings,
        List<X509Certificate2> platformEmbedded,
        CmsVerificationOptions options,
        TimestampVerificationBudget? timestampBudget,
        CertificateUsagePurpose signerCertificatePurpose) {
        var findings = new List<SecurityFinding>();
        int embeddedSignerIndex = -1;
        BcX509Certificate? bcSigner = null;
        for (int index = 0; index < embedded.Count; index++) {
            if (!signer.SignerID.Match(embedded[index])) continue;
            embeddedSignerIndex = index;
            bcSigner = embedded[index];
            break;
        }
        bcSigner ??= FindExtraCertificate(signer.SignerID, options.CertificateValidation.ExtraCertificates);

        if (bcSigner == null) {
            findings.Add(new SecurityFinding(
                SecurityFindingSeverity.Error,
                "SignerCertificateMissing",
                "No certificate matching the CMS signer identifier was supplied or embedded.",
                signerIndex));
            return CreateMissingCertificateResult(signer, signerIndex, findings);
        }

        byte[] encodedSigner = embeddedSignerIndex >= 0
            ? embeddedEncodings[embeddedSignerIndex]
            : bcSigner.GetEncoded();
        X509Certificate2? ownedPlatformSigner = null;
        X509Certificate2 platformSigner = embeddedSignerIndex >= 0
            ? platformEmbedded[embeddedSignerIndex]
            : ownedPlatformSigner = PlatformCertificateLoader.Load(encodedSigner);
        try {
            SecurityValidationStatus digestStatus = ValidateDigest(signer, content, signerIndex, findings);
            SecurityValidationStatus signatureStatus;
            if (content == null) {
                signatureStatus = SecurityValidationStatus.Indeterminate;
            } else {
                try {
                    signatureStatus = VerifySignature(signer, bcSigner, platformSigner, content, useRsaFastPath)
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

            DateTimeOffset? signingTime = ReadSigningTime(signer.SignedAttributes, signerIndex, findings);
            IReadOnlyList<Rfc3161TimestampVerificationResult> timestamps = options.ValidateTimestamps
                ? VerifyTimestamps(signer, options, timestampBudget!, signerIndex, findings)
                : Array.Empty<Rfc3161TimestampVerificationResult>();
            SecurityValidationStatus timestampStatus = options.ValidateTimestamps
                ? AggregateTimestampStatus(timestamps)
                : SecurityValidationStatus.NotPerformed;
            DateTimeOffset? timestampTime = options.ValidateTimestamps
                ? FindLatestValidTimestamp(timestamps)
                : null;
            CertificateValidationOptions signerOptions = ResolveSignerCertificateValidation(
                options.CertificateValidation, timestamps);
            CertificateValidationResult certificateValidation = CertificateChainValidator.Validate(
                platformSigner,
                platformEmbedded,
                signerOptions,
                findings,
                "CMS signer",
                signerCertificatePurpose,
                signerIndex);

            return new CmsSignerVerificationResult(
                signerIndex,
                signatureStatus,
                digestStatus,
                certificateValidation,
                timestampStatus,
                encodedSigner,
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
        } finally {
            ownedPlatformSigner?.Dispose();
        }
    }

    private static CertificateUsagePurpose MapCertificatePurpose(CertificateValidationPurpose purpose) => purpose switch {
        CertificateValidationPurpose.TimestampAuthority => CertificateUsagePurpose.TimestampAuthority,
        CertificateValidationPurpose.EmailSigning => CertificateUsagePurpose.EmailSigner,
        _ => CertificateUsagePurpose.DocumentSigner
    };

    internal static CertificateValidationOptions ResolveSignerCertificateValidation(
        CertificateValidationOptions source,
        IReadOnlyList<Rfc3161TimestampVerificationResult> timestamps) {
        if (source.VerificationTime.HasValue || timestamps.Count == 0) return source;

        DateTime? verificationTime = null;
        for (int index = 0; index < timestamps.Count; index++) {
            Rfc3161TimestampVerificationResult timestamp = timestamps[index];
            if (timestamp.Status != SecurityValidationStatus.Valid || !timestamp.Timestamp.HasValue) continue;
            DateTime candidate = timestamp.Timestamp.Value.UtcDateTime;
            if (!verificationTime.HasValue || candidate < verificationTime.Value) verificationTime = candidate;
        }
        if (!verificationTime.HasValue) return source;

        var result = new CertificateValidationOptions {
            ValidateChain = source.ValidateChain,
            RevocationMode = source.RevocationMode,
            RevocationFlag = source.RevocationFlag,
            VerificationFlags = source.VerificationFlags,
            DisableCertificateDownloads = source.DisableCertificateDownloads,
            VerificationTime = verificationTime,
            UrlRetrievalTimeout = source.UrlRetrievalTimeout,
            ChainEvaluator = source.ChainEvaluator
        };
        result.ExtraCertificates.AddRange(source.ExtraCertificates);
        return result;
    }

    private static AuthenticodeIndirectDataInfo? TryReadAuthenticodeIndirectData(
        string? contentTypeOid,
        byte[]? content,
        List<SecurityFinding> findings) {
        const string SpcIndirectDataContentOid = "1.3.6.1.4.1.311.2.1.4";
        if (!string.Equals(contentTypeOid, SpcIndirectDataContentOid, StringComparison.Ordinal) || content == null) {
            return null;
        }
        try {
            Asn1Encodable dataValue;
            Asn1Encodable digestValue;
            using (var input = new Asn1InputStream(content)) {
                Asn1Object? first = input.ReadObject();
                Asn1Object? second = input.ReadObject();
                Asn1Object? trailing = input.ReadObject();
                if (first is Asn1Sequence wrapped && second == null && wrapped.Count == 2) {
                    dataValue = wrapped[0];
                    digestValue = wrapped[1];
                } else if (first is Asn1Sequence && second != null && trailing == null) {
                    dataValue = first;
                    digestValue = second;
                } else {
                    throw new InvalidDataException("SPC indirect data must contain exactly two values.");
                }
            }
            Asn1Sequence data = Asn1Sequence.GetInstance(dataValue);
            if (data.Count == 0) throw new InvalidDataException("SPC indirect data has no subject type.");
            string subjectType = DerObjectIdentifier.GetInstance(data[0]).Id;
            Org.BouncyCastle.Asn1.X509.DigestInfo digestInfo =
                Org.BouncyCastle.Asn1.X509.DigestInfo.GetInstance(digestValue);
            byte[] digest = digestInfo.Digest.GetOctets();
            if (string.Equals(subjectType, "1.3.6.1.4.1.311.2.1.31", StringComparison.Ordinal)) {
                ValidateVbaV2Descriptor(data);
                if (!OfficeVbaSignatureEncoding.TryExtractV2SourceHash(digest,
                        digestInfo.DigestAlgorithm.Algorithm.Id, out _, out digest, out string v2Detail)) {
                    throw new InvalidDataException(v2Detail);
                }
            }
            if (digest.Length == 0 || digest.Length > 1024) {
                throw new InvalidDataException("The signed subject digest length is invalid.");
            }
            return new AuthenticodeIndirectDataInfo(digestInfo.DigestAlgorithm.Algorithm.Id, digest);
        } catch (Exception exception) when (IsValidationException(exception)) {
            findings.Add(new SecurityFinding(
                SecurityFindingSeverity.Error,
                "AuthenticodeIndirectDataMalformed",
                "The Authenticode SPC indirect-data content could not be decoded: " + exception.Message));
            return null;
        }
    }

    private static void ValidateVbaV2Descriptor(Asn1Sequence data) {
        if (data.Count != 2) {
            throw new InvalidDataException("The VBA V2 indirect-data descriptor is missing or duplicated.");
        }
        Asn1TaggedObject tagged = Asn1TaggedObject.GetInstance(data[1]);
        if (tagged.TagNo != 0 || !tagged.IsExplicit()) {
            throw new InvalidDataException("The VBA V2 indirect-data descriptor tag is invalid.");
        }
        byte[] descriptor = Asn1OctetString.GetInstance(tagged.GetExplicitBaseObject()).GetOctets();
        if (descriptor.Length != 12 || ReadInt32LittleEndian(descriptor, 0) != 12 ||
            ReadInt32LittleEndian(descriptor, 4) != 1 || ReadInt32LittleEndian(descriptor, 8) != 1) {
            throw new InvalidDataException("The VBA V2 signature-format descriptor is invalid.");
        }
    }

    private static int ReadInt32LittleEndian(byte[] bytes, int offset) =>
        bytes[offset] | bytes[offset + 1] << 8 | bytes[offset + 2] << 16 | bytes[offset + 3] << 24;

    private static SecurityValidationStatus ValidateDigest(
        SignerInformation signer,
        byte[]? content,
        int signerIndex,
        List<SecurityFinding> findings) {
        if (content == null) return SecurityValidationStatus.Indeterminate;
        AttributeTable? signedAttributes = signer.SignedAttributes;
        if (signedAttributes == null) return SecurityValidationStatus.Valid;

        Org.BouncyCastle.Asn1.Cms.Attribute? digestAttribute = null;
        bool duplicateDigest = false;
        foreach (Org.BouncyCastle.Asn1.Cms.Attribute attribute in signedAttributes) {
            if (!attribute.AttrType.Equals(CmsAttributes.MessageDigest)) continue;
            if (digestAttribute != null) {
                duplicateDigest = true;
                break;
            }
            digestAttribute = attribute;
        }
        if (duplicateDigest || digestAttribute == null || digestAttribute.AttrValues.Count != 1 ||
            digestAttribute.AttrValues[0] is not Asn1OctetString encodedDigest) {
            findings.Add(new SecurityFinding(
                SecurityFindingSeverity.Error,
                digestAttribute == null ? "CmsMessageDigestMissing" : "CmsMessageDigestInvalid",
                digestAttribute == null
                    ? "Signed attributes must contain a message-digest value."
                    : "Signed attributes must contain exactly one well-formed message-digest value.",
                signerIndex));
            return SecurityValidationStatus.Invalid;
        }

        try {
            byte[] calculated = CalculateDigest(signer.DigestAlgorithmID.Algorithm.Id, content);
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

    private static bool VerifySignature(
        SignerInformation signer,
        BcX509Certificate bcSigner,
        X509Certificate2 platformSigner,
        byte[] content,
        bool useRsaFastPath) {
        if (TryVerifyRsaSignature(signer, platformSigner, content, useRsaFastPath, out bool valid)) return valid;
        return signer.Verify(bcSigner);
    }

    private static bool TryVerifyRsaSignature(
        SignerInformation signer,
        X509Certificate2 certificate,
        byte[] content,
        bool useRsaFastPath,
        out bool valid) {
        valid = false;
        if (!useRsaFastPath ||
            !TryGetHashAlgorithm(signer.DigestAlgorithmID.Algorithm.Id, out HashAlgorithmName digestAlgorithm)) {
            return false;
        }

        using RSA? rsa = certificate.GetRSAPublicKey();
        if (rsa == null) return true;
        byte[] signedBytes = signer.SignedAttributes == null
            ? content
            : signer.GetEncodedSignedAttributes();
        valid = rsa.VerifyData(
            signedBytes,
            signer.GetSignature(),
            digestAlgorithm,
            RSASignaturePadding.Pkcs1);
        return true;
    }

    private static bool CanUseRsaFastPath(SignerInformation signer) =>
        TryGetHashAlgorithm(signer.DigestAlgorithmID.Algorithm.Id, out HashAlgorithmName digestAlgorithm) &&
        IsRsaPkcs1SignatureAlgorithm(signer.SignatureAlgorithm.Algorithm.Id, digestAlgorithm) &&
        HasStandardSignedAttributes(signer);

    private static bool HasStandardSignedAttributes(SignerInformation signer) {
        AttributeTable? attributes = signer.SignedAttributes;
        if (attributes == null) return true;
        if (signer.IsCounterSignature) return false;

        Org.BouncyCastle.Asn1.Cms.Attribute? contentTypeAttribute = null;
        Org.BouncyCastle.Asn1.Cms.Attribute? messageDigestAttribute = null;
        Org.BouncyCastle.Asn1.Cms.Attribute? protectionAttribute = null;
        foreach (Org.BouncyCastle.Asn1.Cms.Attribute attribute in attributes) {
            if (attribute.AttrType.Equals(CmsAttributes.CounterSignature)) return false;
            if (attribute.AttrType.Equals(CmsAttributes.ContentType)) {
                if (contentTypeAttribute != null) return false;
                contentTypeAttribute = attribute;
            } else if (attribute.AttrType.Equals(CmsAttributes.MessageDigest)) {
                if (messageDigestAttribute != null) return false;
                messageDigestAttribute = attribute;
            } else if (attribute.AttrType.Equals(CmsAttributes.CmsAlgorithmProtect)) {
                if (protectionAttribute != null) return false;
                protectionAttribute = attribute;
            }
        }

        if (contentTypeAttribute == null || contentTypeAttribute.AttrValues.Count != 1 ||
            contentTypeAttribute.AttrValues[0] is not DerObjectIdentifier contentType ||
            !contentType.Equals(signer.ContentType)) {
            return false;
        }

        if (messageDigestAttribute == null || messageDigestAttribute.AttrValues.Count != 1 ||
            messageDigestAttribute.AttrValues[0] is not Asn1OctetString) {
            return false;
        }

        if (protectionAttribute == null) return true;
        if (protectionAttribute.AttrValues.Count != 1) return false;

        try {
            CmsAlgorithmProtection protection = CmsAlgorithmProtection.GetInstance(protectionAttribute.AttrValues[0]);
            return protection.MacAlgorithm == null &&
                   protection.DigestAlgorithm.Equals(signer.DigestAlgorithmID) &&
                   protection.SignatureAlgorithm != null &&
                   protection.SignatureAlgorithm.Equals(signer.SignatureAlgorithm);
        } catch (Exception exception) when (IsValidationException(exception)) {
            return false;
        }
    }

    private static byte[] CalculateDigest(string digestAlgorithmOid, byte[] content) {
        if (!TryGetHashAlgorithm(digestAlgorithmOid, out HashAlgorithmName algorithm)) {
            return DigestUtilities.CalculateDigest(digestAlgorithmOid, content);
        }

        using HashAlgorithm hash = algorithm == HashAlgorithmName.SHA256
            ? SHA256.Create()
            : algorithm == HashAlgorithmName.SHA384
                ? SHA384.Create()
                : algorithm == HashAlgorithmName.SHA512
                    ? SHA512.Create()
#pragma warning disable CA5350 // Verification must support caller-supplied legacy SHA-1 CMS signatures.
                    : SHA1.Create();
#pragma warning restore CA5350
        return hash.ComputeHash(content);
    }

    private static bool TryGetHashAlgorithm(string digestAlgorithmOid, out HashAlgorithmName algorithm) {
        switch (digestAlgorithmOid) {
            case "1.3.14.3.2.26":
                algorithm = HashAlgorithmName.SHA1;
                return true;
            case "2.16.840.1.101.3.4.2.1":
                algorithm = HashAlgorithmName.SHA256;
                return true;
            case "2.16.840.1.101.3.4.2.2":
                algorithm = HashAlgorithmName.SHA384;
                return true;
            case "2.16.840.1.101.3.4.2.3":
                algorithm = HashAlgorithmName.SHA512;
                return true;
            default:
                algorithm = default;
                return false;
        }
    }

    private static bool IsRsaPkcs1SignatureAlgorithm(string signatureAlgorithmOid, HashAlgorithmName digestAlgorithm) {
        if (signatureAlgorithmOid == "1.2.840.113549.1.1.1") return true;
        if (signatureAlgorithmOid == "1.2.840.113549.1.1.5") return digestAlgorithm == HashAlgorithmName.SHA1;
        if (signatureAlgorithmOid == "1.2.840.113549.1.1.11") return digestAlgorithm == HashAlgorithmName.SHA256;
        if (signatureAlgorithmOid == "1.2.840.113549.1.1.12") return digestAlgorithm == HashAlgorithmName.SHA384;
        if (signatureAlgorithmOid == "1.2.840.113549.1.1.13") return digestAlgorithm == HashAlgorithmName.SHA512;
        return false;
    }

    private static IReadOnlyList<Rfc3161TimestampVerificationResult> VerifyTimestamps(
        SignerInformation signer,
        CmsVerificationOptions options,
        TimestampVerificationBudget budget,
        int signerIndex,
        List<SecurityFinding> findings) {
        AttributeTable? unsignedAttributes = signer.UnsignedAttributes;
        if (unsignedAttributes == null) return Array.Empty<Rfc3161TimestampVerificationResult>();
        var results = new List<Rfc3161TimestampVerificationResult>();
        foreach (Org.BouncyCastle.Asn1.Cms.Attribute attribute in unsignedAttributes) {
            if (!attribute.AttrType.Equals(PkcsObjectIdentifiers.IdAASignatureTimeStampToken)) continue;
            for (int index = 0; index < attribute.AttrValues.Count; index++) {
                if (!budget.TryReserveToken(out string? limitCode, out string? limitMessage)) {
                    results.Add(CreateTimestampLimitResult(limitCode!, limitMessage!, signerIndex, findings));
                    return results;
                }
                long encodingLimit = budget.GetRemainingEncodingLimit(out limitCode, out limitMessage);
                if (encodingLimit <= 0) {
                    results.Add(CreateTimestampLimitResult(limitCode!, limitMessage!, signerIndex, findings));
                    return results;
                }
                byte[] encoded;
                try {
                    using var encodedToken = new BoundedMemoryStream(encodingLimit);
                    attribute.AttrValues[index].EncodeTo(encodedToken);
                    encoded = encodedToken.ToArray();
                } catch (SecurityContentLimitExceededException) {
                    results.Add(CreateTimestampLimitResult(limitCode!, limitMessage!, signerIndex, findings));
                    return results;
                }
                if (!budget.TryReserveBytes(encoded.LongLength, out limitCode, out limitMessage)) {
                    results.Add(CreateTimestampLimitResult(limitCode!, limitMessage!, signerIndex, findings));
                    return results;
                }
                Rfc3161TimestampVerificationResult result = Rfc3161TimestampVerifier.Verify(
                    encoded,
                    signer.GetSignature(),
                    options.CertificateValidation,
                    options.MaxTimestampTokenBytes,
                    options.MaxCertificates);
                results.Add(result);
                foreach (SecurityFinding finding in result.Findings) {
                    findings.Add(new SecurityFinding(finding.Severity, finding.Code, finding.Message, signerIndex));
                }
            }
        }

        return results;
    }

    internal static void ValidateOptions(CmsVerificationOptions options) {
#if NETSTANDARD2_0 || NET472
        if (options == null) throw new ArgumentNullException(nameof(options));
#else
        ArgumentNullException.ThrowIfNull(options);
#endif
        SecurityLimits.EnsureBufferWithinLimit(Array.Empty<byte>(), options.MaxEncodedBytes,
            nameof(options.MaxEncodedBytes));
        SecurityLimits.EnsureBufferWithinLimit(Array.Empty<byte>(), options.MaxContentBytes,
            nameof(options.MaxContentBytes));
        SecurityLimits.EnsureCountWithinLimit(0, options.MaxSigners, nameof(options.MaxSigners));
        SecurityLimits.EnsureCountWithinLimit(0, options.MaxCertificates, nameof(options.MaxCertificates));
        SecurityLimits.EnsureCountWithinLimit(0, options.MaxTimestampTokens, nameof(options.MaxTimestampTokens));
        SecurityLimits.EnsureBufferWithinLimit(Array.Empty<byte>(), options.MaxTimestampTokenBytes,
            nameof(options.MaxTimestampTokenBytes));
        SecurityLimits.EnsureBufferWithinLimit(Array.Empty<byte>(), options.MaxTotalTimestampBytes,
            nameof(options.MaxTotalTimestampBytes));
    }

    private static Rfc3161TimestampVerificationResult CreateTimestampLimitResult(
        string code,
        string message,
        int signerIndex,
        List<SecurityFinding> signerFindings) {
        var finding = new SecurityFinding(SecurityFindingSeverity.Error, code, message, signerIndex);
        signerFindings.Add(finding);
        return new Rfc3161TimestampVerificationResult(
            SecurityValidationStatus.Invalid,
            null,
            null,
            null,
            null,
            new CertificateValidationResult(
                SecurityValidationStatus.Indeterminate,
                SecurityValidationStatus.NotPerformed,
                Array.Empty<string>()),
            new[] { finding });
    }

    internal sealed class TimestampVerificationBudget {
        private readonly int _maximumTokens;
        private readonly long _maximumTokenBytes;
        private readonly long _maximumTotalBytes;
        private int _tokens;
        private long _totalBytes;

        internal TimestampVerificationBudget(CmsVerificationOptions options) {
            _maximumTokens = options.MaxTimestampTokens;
            _maximumTokenBytes = options.MaxTimestampTokenBytes;
            _maximumTotalBytes = options.MaxTotalTimestampBytes;
        }

        internal bool TryReserveToken(out string? code, out string? message) {
            if (_tokens >= _maximumTokens) {
                code = "CmsTimestampCountLimitExceeded";
                message = $"CMS timestamp tokens exceed the configured aggregate limit of {_maximumTokens}.";
                return false;
            }
            _tokens++;
            code = null;
            message = null;
            return true;
        }

        internal bool TryReserveBytes(long encodedBytes, out string? code, out string? message) {
            if (encodedBytes > _maximumTokenBytes) {
                code = "CmsTimestampSizeLimitExceeded";
                message = $"An encoded CMS timestamp token exceeds the configured limit of {_maximumTokenBytes} bytes.";
                return false;
            }
            if (encodedBytes > _maximumTotalBytes - _totalBytes) {
                code = "CmsTimestampTotalSizeLimitExceeded";
                message = $"Encoded CMS timestamp tokens exceed the configured aggregate limit of {_maximumTotalBytes} bytes.";
                return false;
            }
            _totalBytes += encodedBytes;
            code = null;
            message = null;
            return true;
        }

        internal long GetRemainingEncodingLimit(out string code, out string message) {
            long remainingTotal = _maximumTotalBytes - _totalBytes;
            if (_maximumTokenBytes <= remainingTotal) {
                code = "CmsTimestampSizeLimitExceeded";
                message = $"An encoded CMS timestamp token exceeds the configured limit of {_maximumTokenBytes} bytes.";
                return _maximumTokenBytes;
            }

            code = "CmsTimestampTotalSizeLimitExceeded";
            message = $"Encoded CMS timestamp tokens exceed the configured aggregate limit of {_maximumTotalBytes} bytes.";
            return remainingTotal;
        }
    }

    private static DateTimeOffset? ReadSigningTime(
        AttributeTable? signedAttributes,
        int signerIndex,
        List<SecurityFinding> findings) {
        if (signedAttributes == null) return null;
        Org.BouncyCastle.Asn1.Cms.Attribute? signingTimeAttribute = null;
        foreach (Org.BouncyCastle.Asn1.Cms.Attribute attribute in signedAttributes) {
            if (!attribute.AttrType.Equals(CmsAttributes.SigningTime)) continue;
            if (signingTimeAttribute == null) {
                signingTimeAttribute = attribute;
                continue;
            }
            findings.Add(new SecurityFinding(
                SecurityFindingSeverity.Warning,
                "CmsSigningTimeInvalid",
                "The signing-time attribute is duplicated or malformed.",
                signerIndex));
            return null;
        }
        if (signingTimeAttribute == null) return null;
        if (signingTimeAttribute.AttrValues.Count != 1) {
            findings.Add(new SecurityFinding(
                SecurityFindingSeverity.Warning,
                "CmsSigningTimeInvalid",
                "The signing-time attribute is duplicated or malformed.",
                signerIndex));
            return null;
        }
        try {
            DateTime value = Org.BouncyCastle.Asn1.Cms.Time.GetInstance(signingTimeAttribute.AttrValues[0]).ToDateTime();
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

    internal static BcX509Certificate? FindExtraCertificate(
        ISelector<BcX509Certificate> selector,
        X509Certificate2Collection extraCertificates) {
        foreach (X509Certificate2 certificate in extraCertificates) {
            BcX509Certificate candidate = DotNetUtilities.FromX509Certificate(certificate);
            if (selector.Match(candidate)) return candidate;
        }
        return null;
    }

    private static List<X509Certificate2> CreatePlatformCertificates(
        List<BcX509Certificate> certificates,
        List<SecurityFinding> findings,
        out List<byte[]> encodings) {
        var result = new List<X509Certificate2>(certificates.Count);
        encodings = new List<byte[]>(certificates.Count);
        try {
            foreach (BcX509Certificate certificate in certificates) {
                byte[] encoded = certificate.GetEncoded();
                encodings.Add(encoded);
                result.Add(PlatformCertificateLoader.Load(encoded));
            }
            return result;
        } catch (Exception exception) when (IsValidationException(exception)) {
            foreach (X509Certificate2 certificate in result) certificate.Dispose();
            encodings.Clear();
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

    private static DateTimeOffset? FindLatestValidTimestamp(
        IReadOnlyList<Rfc3161TimestampVerificationResult> timestamps) {
        DateTimeOffset? latest = null;
        for (int index = 0; index < timestamps.Count; index++) {
            Rfc3161TimestampVerificationResult result = timestamps[index];
            if (result.Status == SecurityValidationStatus.Valid && result.Timestamp.HasValue &&
                (!latest.HasValue || result.Timestamp.Value > latest.Value)) {
                latest = result.Timestamp.Value;
            }
        }
        return latest;
    }

    private static bool IsValidationException(Exception exception) =>
        exception is not OutOfMemoryException &&
        exception is not StackOverflowException &&
        exception is not AccessViolationException;
}
