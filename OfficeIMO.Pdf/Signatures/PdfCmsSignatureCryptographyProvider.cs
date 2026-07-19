using System.Security.Cryptography;
using OfficeIMO.Security;

namespace OfficeIMO.Pdf;

/// <summary>Maps shared CMS, RFC 3161, and X.509 validation into PDF signature reports.</summary>
public sealed class PdfCmsSignatureCryptographyProvider : IPdfSignatureCryptographyProvider {
    private readonly CmsVerificationOptions _options;

    /// <summary>Creates a PDF adapter over the shared OfficeIMO security policy.</summary>
    public PdfCmsSignatureCryptographyProvider(CmsVerificationOptions? options = null) {
        _options = options ?? new CmsVerificationOptions();
    }

    /// <inheritdoc />
    public string Name => "OfficeIMO.Pdf CMS";

    /// <inheritdoc />
    public PdfSignatureCryptographicResult Verify(PdfSignatureCryptographyInput input) {
#if NETSTANDARD2_0 || NET472
        if (input == null) throw new ArgumentNullException(nameof(input));
#else
        ArgumentNullException.ThrowIfNull(input);
#endif
        byte[] encoded;
        try {
            encoded = SecurityEncoding.NormalizeSingleAsn1Object(
                input.SignatureContents,
                allowTrailingZeroPadding: true,
                _options.MaxEncodedBytes);
        } catch (Exception exception) when (IsValidationException(exception)) {
            return InvalidContainer("CmsSignatureInvalid", "The signature container could not be decoded: " + exception.Message);
        }

        return input.Signature.IsDocumentTimestamp
            ? VerifyDocumentTimestamp(input.SignedContent, encoded)
            : VerifyCmsSignature(input, encoded);
    }

    private PdfSignatureCryptographicResult VerifyCmsSignature(
        PdfSignatureCryptographyInput input,
        byte[] encoded) {
        bool encapsulatedSha1 = string.Equals(input.Signature.SubFilter, "adbe.pkcs7.sha1", StringComparison.Ordinal);
        CmsVerificationResult security = encapsulatedSha1
            ? CmsSignedDataVerifier.Verify(encoded, _options)
            : CmsSignedDataVerifier.VerifyDetached(encoded, input.SignedContent, _options);
        var findings = MapFindings(security.Findings, security.Signers);

        if (!security.Parsed || security.Signers.Count == 0) {
            return InvalidContainer(findings);
        }
        if (!encapsulatedSha1 && !security.IsDetached) {
            findings.Add(Finding(
                PdfDiagnosticSeverity.Error,
                "CmsDetachedContentExpected",
                "Detached PDF signatures must not carry encapsulated CMS content."));
            return InvalidContainer(findings);
        }

        PdfCryptographicValidationStatus digestStatus = Aggregate(
            security.Signers.Select(static signer => signer.DigestStatus));
        PdfCryptographicValidationStatus signatureStatus = Aggregate(
            security.Signers.Select(static signer => signer.SignatureStatus));
        if (findings.Any(static finding => finding.Code == "CmsSignerMissing")) {
            signatureStatus = PdfCryptographicValidationStatus.Invalid;
        }
        if (encapsulatedSha1) {
            bool sha1Matches = security.EncapsulatedContent != null &&
                FixedTimeEquals(ComputeSha1(input.SignedContent), security.EncapsulatedContent);
            if (!sha1Matches) {
                digestStatus = PdfCryptographicValidationStatus.Invalid;
                signatureStatus = PdfCryptographicValidationStatus.Invalid;
                findings.Add(Finding(
                    PdfDiagnosticSeverity.Error,
                    "CmsDigestMismatch",
                    "The encapsulated SHA-1 value does not match the PDF signed byte ranges."));
            }
        }

        if (digestStatus == PdfCryptographicValidationStatus.Invalid &&
            findings.All(static finding => finding.Code != "CmsDigestMismatch")) {
            findings.Add(Finding(
                PdfDiagnosticSeverity.Error,
                "CmsDigestMismatch",
                "The CMS message digest does not match the PDF signed byte ranges."));
        }

        PdfCryptographicValidationStatus chainStatus = Aggregate(
            security.Signers.Select(static signer => signer.CertificateValidation.ChainStatus));
        PdfCryptographicValidationStatus revocationStatus = Aggregate(
            security.Signers.Select(static signer => signer.CertificateValidation.RevocationStatus));
        PdfCryptographicValidationStatus timestampStatus = AggregateTimestampStatus(security.Signers);
        if (timestampStatus == PdfCryptographicValidationStatus.Invalid) {
            findings.Add(Finding(
                PdfDiagnosticSeverity.Error,
                "SignatureTimestampInvalid",
                "At least one CMS signature timestamp or its TSA trust policy did not validate."));
        }

        CmsSignerVerificationResult first = security.Signers[0];
        return new PdfSignatureCryptographicResult(
            Name,
            signatureStatus,
            digestStatus,
            chainStatus,
            revocationStatus,
            timestampStatus,
            first.Subject,
            first.Issuer,
            first.SerialNumber,
            first.Thumbprint,
            first.SigningTime,
            first.TimestampTime,
            findings);
    }

    private PdfSignatureCryptographicResult VerifyDocumentTimestamp(byte[] signedContent, byte[] encoded) {
        if (!_options.ValidateTimestamps) {
            return new PdfSignatureCryptographicResult(
                Name,
                PdfCryptographicValidationStatus.NotPerformed,
                PdfCryptographicValidationStatus.NotPerformed,
                PdfCryptographicValidationStatus.NotPerformed,
                PdfCryptographicValidationStatus.NotPerformed,
                PdfCryptographicValidationStatus.NotPerformed,
                findings: new[] {
                    Finding(PdfDiagnosticSeverity.Info, "TimestampValidationDisabled", "RFC 3161 validation was disabled by provider policy.")
                });
        }

        Rfc3161TimestampVerificationResult timestamp = Rfc3161TimestampVerifier.Verify(
            encoded,
            signedContent,
            _options.CertificateValidation,
            _options.MaxEncodedBytes,
            _options.MaxCertificates);
        var findings = timestamp.Findings.Select(MapFinding).ToList();
        PdfCryptographicValidationStatus cryptoStatus = MapStatus(timestamp.Status);
        PdfCryptographicValidationStatus chainStatus = MapStatus(timestamp.CertificateValidation.ChainStatus);
        PdfCryptographicValidationStatus timestampStatus = chainStatus == PdfCryptographicValidationStatus.Invalid
            ? PdfCryptographicValidationStatus.Invalid
            : cryptoStatus;
        if (timestampStatus == PdfCryptographicValidationStatus.Invalid) {
            findings.Add(Finding(
                PdfDiagnosticSeverity.Error,
                "TimestampMessageImprintInvalid",
                "The RFC 3161 token, TSA trust policy, or message imprint did not validate against the PDF signed byte ranges."));
        }

        return new PdfSignatureCryptographicResult(
            Name,
            cryptoStatus,
            cryptoStatus,
            chainStatus,
            MapStatus(timestamp.CertificateValidation.RevocationStatus),
            timestampStatus,
            timestampTime: timestamp.Timestamp,
            findings: findings);
    }

    private PdfSignatureCryptographicResult InvalidContainer(string code, string message) =>
        InvalidContainer(new List<PdfSignatureCryptographicFinding> { Finding(PdfDiagnosticSeverity.Error, code, message) });

    private PdfSignatureCryptographicResult InvalidContainer(List<PdfSignatureCryptographicFinding> findings) =>
        new PdfSignatureCryptographicResult(
            Name,
            PdfCryptographicValidationStatus.Invalid,
            PdfCryptographicValidationStatus.Invalid,
            PdfCryptographicValidationStatus.NotPerformed,
            PdfCryptographicValidationStatus.NotPerformed,
            PdfCryptographicValidationStatus.NotPerformed,
            findings: findings);

    private static List<PdfSignatureCryptographicFinding> MapFindings(
        IReadOnlyList<SecurityFinding> containerFindings,
        IReadOnlyList<CmsSignerVerificationResult> signers) {
        var result = containerFindings.Select(MapFinding).ToList();
        foreach (CmsSignerVerificationResult signer in signers) {
            result.AddRange(signer.Findings.Select(MapFinding));
        }
        if (result.Any(static finding => finding.Code == "SignerCertificateMissing")) {
            result.Add(Finding(PdfDiagnosticSeverity.Error, "CmsSignerMissing", "The CMS signer certificate was not embedded or supplied by caller policy."));
        }
        return result;
    }

    private static PdfCryptographicValidationStatus AggregateTimestampStatus(
        IReadOnlyList<CmsSignerVerificationResult> signers) {
        var statuses = new List<SecurityValidationStatus>(signers.Count);
        foreach (CmsSignerVerificationResult signer in signers) {
            SecurityValidationStatus status = signer.TimestampStatus;
            if (signer.TimestampTokens.Any(static token =>
                    token.CertificateValidation.ChainStatus == SecurityValidationStatus.Invalid)) {
                status = SecurityValidationStatus.Invalid;
            }
            statuses.Add(status);
        }
        return Aggregate(statuses);
    }

    private static PdfCryptographicValidationStatus Aggregate(IEnumerable<SecurityValidationStatus> statuses) {
        SecurityValidationStatus[] values = statuses.ToArray();
        if (values.Length == 0) return PdfCryptographicValidationStatus.NotPerformed;
        if (values.Any(static value => value == SecurityValidationStatus.Invalid)) return PdfCryptographicValidationStatus.Invalid;
        if (values.Any(static value => value == SecurityValidationStatus.Indeterminate)) return PdfCryptographicValidationStatus.Indeterminate;
        if (values.Any(static value => value == SecurityValidationStatus.Valid)) return PdfCryptographicValidationStatus.Valid;
        return PdfCryptographicValidationStatus.NotPerformed;
    }

    private static PdfCryptographicValidationStatus MapStatus(SecurityValidationStatus status) => status switch {
        SecurityValidationStatus.NotPerformed => PdfCryptographicValidationStatus.NotPerformed,
        SecurityValidationStatus.Valid => PdfCryptographicValidationStatus.Valid,
        SecurityValidationStatus.Invalid => PdfCryptographicValidationStatus.Invalid,
        SecurityValidationStatus.Indeterminate => PdfCryptographicValidationStatus.Indeterminate,
        _ => PdfCryptographicValidationStatus.Error
    };

    private static PdfSignatureCryptographicFinding MapFinding(SecurityFinding finding) =>
        Finding(finding.Severity switch {
            SecurityFindingSeverity.Info => PdfDiagnosticSeverity.Info,
            SecurityFindingSeverity.Warning => PdfDiagnosticSeverity.Warning,
            SecurityFindingSeverity.Error => PdfDiagnosticSeverity.Error,
            _ => PdfDiagnosticSeverity.Warning
        }, finding.Code, finding.Message);

    private static PdfSignatureCryptographicFinding Finding(
        PdfDiagnosticSeverity severity,
        string code,
        string message) => new PdfSignatureCryptographicFinding(severity, code, message);

#pragma warning disable CA1850, CA5350 // Cross-target SHA-1 is required only for legacy adbe.pkcs7.sha1 validation.
    private static byte[] ComputeSha1(byte[] value) {
        using SHA1 sha1 = SHA1.Create();
        return sha1.ComputeHash(value);
    }
#pragma warning restore CA1850, CA5350

    private static bool FixedTimeEquals(byte[] left, byte[] right) {
        if (left.Length != right.Length) return false;
        int difference = 0;
        for (int index = 0; index < left.Length; index++) difference |= left[index] ^ right[index];
        return difference == 0;
    }

    private static bool IsValidationException(Exception exception) =>
        exception is not OutOfMemoryException &&
        exception is not StackOverflowException &&
        exception is not AccessViolationException;
}
