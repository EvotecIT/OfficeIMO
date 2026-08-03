using System.IO;
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
                        authenticodeIndirectData: null,
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
                TimestampVerificationBudget? effectiveTimestampBudget = options.ValidateTimestamps
                    ? timestampBudget ?? new TimestampVerificationBudget(options)
                    : null;
                for (int index = 0; index < signers.Count; index++) {
                    signerResults.Add(VerifySigner(
                        signers[index],
                        index,
                        content,
                        verifiable,
                        embedded,
                        platformEmbedded,
                        options,
                        effectiveTimestampBudget,
                        signerCertificatePurpose));
                }

                AuthenticodeIndirectDataInfo? authenticode = TryReadAuthenticodeIndirectData(
                    verifiable.SignedContentType?.Id, content, containerFindings);
                return new CmsVerificationResult(
                    parsed: true,
                    isDetached,
                    verifiable.SignedContentType?.Id,
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
        CmsSignedData signedData,
        IReadOnlyList<BcX509Certificate> embedded,
        IReadOnlyList<X509Certificate2> platformEmbedded,
        CmsVerificationOptions options,
        TimestampVerificationBudget? timestampBudget,
        CertificateUsagePurpose signerCertificatePurpose) {
        var findings = new List<SecurityFinding>();
        BcX509Certificate? bcSigner = signedData.GetCertificates()
            .EnumerateMatches(signer.SignerID)
            .FirstOrDefault();
        bcSigner ??= FindExtraCertificate(signer.SignerID, options.CertificateValidation.ExtraCertificates);

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

        DateTimeOffset? signingTime = ReadSigningTime(signer.SignedAttributes, signerIndex, findings);
        IReadOnlyList<Rfc3161TimestampVerificationResult> timestamps = options.ValidateTimestamps
            ? VerifyTimestamps(signer, options, timestampBudget!, signerIndex, findings)
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

    private static CertificateUsagePurpose MapCertificatePurpose(CertificateValidationPurpose purpose) => purpose switch {
        CertificateValidationPurpose.TimestampAuthority => CertificateUsagePurpose.TimestampAuthority,
        CertificateValidationPurpose.EmailSigning => CertificateUsagePurpose.EmailSigner,
        _ => CertificateUsagePurpose.DocumentSigner
    };

    internal static CertificateValidationOptions ResolveSignerCertificateValidation(
        CertificateValidationOptions source,
        IReadOnlyList<Rfc3161TimestampVerificationResult> timestamps) {
        DateTime? verificationTime = source.VerificationTime;
        if (verificationTime == null) {
            verificationTime = timestamps
                .Where(static result => result.Status == SecurityValidationStatus.Valid && result.Timestamp.HasValue)
                .Select(static result => (DateTime?)result.Timestamp!.Value.UtcDateTime)
                .OrderBy(static value => value)
                .FirstOrDefault();
        }
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
            Asn1Encodable digestValue;
            using (var input = new Asn1InputStream(content)) {
                Asn1Object? first = input.ReadObject();
                Asn1Object? second = input.ReadObject();
                Asn1Object? trailing = input.ReadObject();
                if (first is Asn1Sequence wrapped && second == null && wrapped.Count == 2) {
                    digestValue = wrapped[1];
                } else if (first is Asn1Sequence && second != null && trailing == null) {
                    digestValue = second;
                } else {
                    throw new InvalidDataException("SPC indirect data must contain exactly two values.");
                }
            }
            Org.BouncyCastle.Asn1.X509.DigestInfo digestInfo =
                Org.BouncyCastle.Asn1.X509.DigestInfo.GetInstance(digestValue);
            byte[] digest = digestInfo.Digest.GetOctets();
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
