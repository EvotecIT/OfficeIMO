using System.Security.Cryptography.X509Certificates;
using Org.BouncyCastle.Security;
using Org.BouncyCastle.Tsp;
using Org.BouncyCastle.Utilities;
using Org.BouncyCastle.X509;
using BcX509Certificate = Org.BouncyCastle.X509.X509Certificate;

namespace OfficeIMO.Security;

/// <summary>Verifies RFC 3161 timestamp tokens against the data they timestamp.</summary>
public static class Rfc3161TimestampVerifier {
    /// <summary>Verifies a timestamp token and its message imprint.</summary>
    public static Rfc3161TimestampVerificationResult Verify(
        byte[] encodedToken,
        byte[] timestampedData,
        CertificateValidationOptions? certificateValidation = null,
        long maxEncodedBytes = 16L * 1024 * 1024,
        int maxCertificates = 64) {
#if NETSTANDARD2_0 || NET472
        if (encodedToken == null) throw new ArgumentNullException(nameof(encodedToken));
        if (timestampedData == null) throw new ArgumentNullException(nameof(timestampedData));
#else
        ArgumentNullException.ThrowIfNull(encodedToken);
        ArgumentNullException.ThrowIfNull(timestampedData);
#endif
        SecurityLimits.EnsureBufferWithinLimit(encodedToken, maxEncodedBytes, nameof(encodedToken));
        var findings = new SecurityFindingCollection();
        CertificateValidationResult emptyValidation = EmptyCertificateValidation();

        try {
            var token = new TimeStampToken(new Org.BouncyCastle.Cms.CmsSignedData(encodedToken));
            TimeStampTokenInfo info = token.TimeStampInfo;
            CertificateValidationOptions effectiveCertificateValidation =
                ResolveCertificateValidation(certificateValidation, info.GenTime);
            byte[] calculatedImprint = DigestUtilities.CalculateDigest(info.MessageImprintAlgOid, timestampedData);
            bool imprintValid = Arrays.FixedTimeEquals(calculatedImprint, info.GetMessageImprintDigest());
            if (!imprintValid) {
                findings.Add(new SecurityFinding(
                    SecurityFindingSeverity.Error,
                    "TimestampImprintMismatch",
                    "The timestamp message imprint does not match the supplied data."));
            }

            List<BcX509Certificate> embedded = token.GetCertificates().EnumerateMatches(null).ToList();
            SecurityLimits.EnsureCountWithinLimit(embedded.Count, maxCertificates, nameof(maxCertificates));
            BcX509Certificate? tsaCertificate = token.GetCertificates()
                .EnumerateMatches(token.SignerID)
                .FirstOrDefault();
            tsaCertificate ??= CmsSignedDataVerifier.FindExtraCertificate(
                token.SignerID,
                effectiveCertificateValidation.ExtraCertificates);
            if (tsaCertificate == null) {
                findings.Add(new SecurityFinding(
                    SecurityFindingSeverity.Error,
                    "TimestampCertificateMissing",
                    "No TSA signing certificate matching the timestamp signer identifier was supplied or embedded."));
                return CreateResult(
                    imprintValid ? SecurityValidationStatus.Indeterminate : SecurityValidationStatus.Invalid,
                    info,
                    null,
                    emptyValidation,
                    findings.Items);
            }

            bool signatureValid;
            try {
                token.Validate(tsaCertificate);
                signatureValid = true;
            } catch (Exception exception) when (IsValidationException(exception)) {
                signatureValid = false;
                findings.Add(new SecurityFinding(
                    SecurityFindingSeverity.Error,
                    "TimestampSignatureInvalid",
                    "The timestamp-token signature or TSA certificate profile is invalid: " + exception.Message));
            }

            using X509Certificate2 platformTsa = PlatformCertificateLoader.Load(tsaCertificate.GetEncoded());
            var platformEmbedded = new List<X509Certificate2>(embedded.Count);
            try {
                foreach (BcX509Certificate certificate in embedded) {
                    platformEmbedded.Add(PlatformCertificateLoader.Load(certificate.GetEncoded()));
                }
                CertificateValidationResult chain = CertificateChainValidator.Validate(
                    platformTsa,
                    platformEmbedded,
                    effectiveCertificateValidation,
                    ref findings,
                    "TSA",
                    CertificateUsagePurpose.TimestampAuthority);
                SecurityValidationStatus status = ResolveTimestampStatus(
                    signatureValid,
                    imprintValid,
                    chain.ChainStatus,
                    chain.RevocationStatus);
                return CreateResult(status, info, tsaCertificate.GetEncoded(), chain, findings.Items);
            } finally {
                foreach (X509Certificate2 certificate in platformEmbedded) certificate.Dispose();
            }
        } catch (Exception exception) when (IsValidationException(exception)) {
            findings.Add(new SecurityFinding(
                SecurityFindingSeverity.Error,
                "TimestampMalformed",
                "The RFC 3161 timestamp token could not be decoded: " + exception.Message));
            return new Rfc3161TimestampVerificationResult(
                SecurityValidationStatus.Invalid,
                null,
                null,
                null,
                null,
                emptyValidation,
                findings.Items);
        }
    }

    private static Rfc3161TimestampVerificationResult CreateResult(
        SecurityValidationStatus status,
        TimeStampTokenInfo info,
        byte[]? certificate,
        CertificateValidationResult certificateValidation,
        IReadOnlyList<SecurityFinding> findings) {
        DateTime utcTime = info.GenTime.Kind == DateTimeKind.Utc
            ? info.GenTime
            : DateTime.SpecifyKind(info.GenTime, DateTimeKind.Utc);
        return new Rfc3161TimestampVerificationResult(
            status,
            new DateTimeOffset(utcTime),
            info.Policy,
            info.MessageImprintAlgOid,
            certificate,
            certificateValidation,
            findings);
    }

    private static CertificateValidationResult EmptyCertificateValidation() =>
        new CertificateValidationResult(
            SecurityValidationStatus.Indeterminate,
            SecurityValidationStatus.NotPerformed,
            Array.Empty<string>());

    internal static SecurityValidationStatus ResolveTimestampStatus(
        bool signatureValid,
        bool imprintValid,
        SecurityValidationStatus certificateStatus,
        SecurityValidationStatus revocationStatus) {
        if (!signatureValid ||
            !imprintValid ||
            certificateStatus == SecurityValidationStatus.Invalid ||
            revocationStatus == SecurityValidationStatus.Invalid) {
            return SecurityValidationStatus.Invalid;
        }
        return certificateStatus == SecurityValidationStatus.Valid &&
               revocationStatus is SecurityValidationStatus.Valid or SecurityValidationStatus.NotPerformed
            ? SecurityValidationStatus.Valid
            : SecurityValidationStatus.Indeterminate;
    }

    private static CertificateValidationOptions ResolveCertificateValidation(
        CertificateValidationOptions? source,
        DateTime generationTime) {
        if (source?.VerificationTime != null) return source;
        var result = new CertificateValidationOptions {
            ValidateChain = source?.ValidateChain ?? true,
            RevocationMode = source?.RevocationMode ?? X509RevocationMode.NoCheck,
            RevocationFlag = source?.RevocationFlag ?? X509RevocationFlag.ExcludeRoot,
            VerificationFlags = source?.VerificationFlags ?? X509VerificationFlags.NoFlag,
            DisableCertificateDownloads = source?.DisableCertificateDownloads ?? true,
            VerificationTime = generationTime.Kind == DateTimeKind.Utc
                ? generationTime
                : DateTime.SpecifyKind(generationTime, DateTimeKind.Utc),
            UrlRetrievalTimeout = source?.UrlRetrievalTimeout ?? TimeSpan.FromSeconds(15),
            ChainEvaluator = source?.ChainEvaluator
        };
        if (source != null) result.ExtraCertificates.AddRange(source.ExtraCertificates);
        return result;
    }

    private static bool IsValidationException(Exception exception) =>
        exception is not OutOfMemoryException &&
        exception is not StackOverflowException &&
        exception is not AccessViolationException;
}
