using System.Security.Cryptography;
using OfficeIMO.Security;

namespace OfficeIMO.Pdf;

/// <summary>Maps shared CMS, RFC 3161, and X.509 validation into PDF signature reports.</summary>
public sealed class PdfCmsSignatureCryptographyProvider : IPdfSignatureCryptographyProvider {
    private readonly IOfficeSecurityProvider _securityProvider;
    private readonly CmsVerificationOptions _options;
    private readonly CertificateRevocationPolicy _revocationPolicy;

    /// <summary>Creates a PDF adapter over the shared OfficeIMO security policy.</summary>
    public PdfCmsSignatureCryptographyProvider(
        IOfficeSecurityProvider securityProvider,
        CmsVerificationOptions? options = null) {
        ArgumentNullException.ThrowIfNull(securityProvider);
        _securityProvider = securityProvider;
        _options = options ?? new CmsVerificationOptions();
    }

    /// <inheritdoc />
    public string Name => "OfficeIMO.Pdf CMS";

    /// <summary>
    /// Where OCSP responses and CRLs come from. Defaults to <see cref="CertificateRevocationPolicy.EmbeddedOnly"/>:
    /// only data carried by the signature and the document's DSS is used and validation never starts network
    /// retrieval. The online policies use <see cref="RevocationFetcher"/>.
    /// </summary>
    public CertificateRevocationPolicy RevocationPolicy {
        get => _revocationPolicy;
        init {
            if (value is < CertificateRevocationPolicy.EmbeddedOnly or > CertificateRevocationPolicy.Online) {
                throw new ArgumentOutOfRangeException(nameof(value), value, "Unknown revocation policy.");
            }
            _revocationPolicy = value;
        }
    }

    /// <summary>
    /// Fetcher used by the online revocation policies, such as <c>CertificateRevocationFetcher</c>. Without one,
    /// certificates that embedded data does not settle are reported as indeterminate.
    /// </summary>
    public ICertificateRevocationFetcher? RevocationFetcher { get; init; }

    /// <summary>
    /// Fetches revocation data for the signer and time-stamp authority paths of one validated signature,
    /// including the certificates of those paths, for embedding in the document's DSS.
    /// </summary>
    internal async Task<CertificateRevocationEvidence> FetchRevocationEvidenceAsync(
        PdfSignatureValidationResult signature,
        ICertificateRevocationFetcher fetcher,
        CancellationToken cancellationToken) {
        if (_securityProvider is not IOfficeRevocationValidator validator) {
            throw new NotSupportedException(_securityProvider.Name + " does not support revocation data retrieval.");
        }
        byte[] contents = signature.Signature.ContentsBytes ??
            throw new InvalidOperationException("The signature has no decoded /Contents bytes.");
        byte[] encoded = _securityProvider.NormalizeAsn1Object(contents, allowTrailingZeroPadding: true, _options.MaxEncodedBytes);
        return await validator.FetchRevocationEvidenceAsync(
            encoded,
            signature.Signature.IsDocumentTimestamp,
            fetcher,
            _options.CertificateValidation.ExtraCertificates,
            cancellationToken).ConfigureAwait(false);
    }

    /// <inheritdoc />
    public PdfSignatureCryptographicResult Verify(PdfSignatureCryptographyInput input) {
#if NETSTANDARD2_0 || NET472
        ArgumentNullException.ThrowIfNull(input);
#else
        ArgumentNullException.ThrowIfNull(input);
#endif
        byte[] encoded;
        try {
            encoded = _securityProvider.NormalizeAsn1Object(
                input.SignatureContents,
                allowTrailingZeroPadding: true,
                _options.MaxEncodedBytes);
        } catch (Exception exception) when (IsValidationException(exception)) {
            return InvalidContainer("cms-signature-invalid", "The signature container could not be decoded: " + exception.Message);
        }

        return input.Signature.IsDocumentTimestamp
            ? VerifyDocumentTimestamp(input, encoded)
            : VerifyCmsSignature(input, encoded);
    }

    private PdfSignatureCryptographicResult VerifyCmsSignature(
        PdfSignatureCryptographyInput input,
        byte[] encoded) {
        bool encapsulatedSha1 = string.Equals(input.Signature.SubFilter, "adbe.pkcs7.sha1", StringComparison.Ordinal);
        CmsVerificationResult security = encapsulatedSha1
            ? _securityProvider.VerifyCms(encoded, _options)
            : _securityProvider.VerifyCmsDetached(encoded, input.SignedContent, _options);
        var findings = MapFindings(security.Findings, security.Signers);

        if (!security.Parsed || security.Signers.Count == 0) {
            return InvalidContainer(findings);
        }
        if (!encapsulatedSha1 && !security.IsDetached) {
            findings.Add(Finding(
                PdfDiagnosticSeverity.Error,
                "cms-detached-content-expected",
                "Detached PDF signatures must not carry encapsulated CMS content."));
            return InvalidContainer(findings);
        }

        PdfCryptographicValidationStatus digestStatus = Aggregate(
            security.Signers.Select(static signer => signer.DigestStatus));
        PdfCryptographicValidationStatus signatureStatus = Aggregate(
            security.Signers.Select(static signer => signer.SignatureStatus));
        if (findings.Any(static finding => finding.Code == "cms-signer-missing")) {
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
                    "cms-digest-mismatch",
                    "The encapsulated SHA-1 value does not match the PDF signed byte ranges."));
            }
        }

        if (digestStatus == PdfCryptographicValidationStatus.Invalid &&
            findings.All(static finding => finding.Code != "cms-digest-mismatch")) {
            findings.Add(Finding(
                PdfDiagnosticSeverity.Error,
                "cms-digest-mismatch",
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
                "signature-timestamp-invalid",
                "At least one CMS signature timestamp or its TSA trust policy did not validate."));
        }

        CmsSignerVerificationResult first = security.Signers[0];
        OfficeRevocationOutcome? revocation = ValidateRevocation(
            input,
            encoded,
            isTimestampToken: false,
            ResolveSignerValidationTime(security.Signers),
            findings);
        if (revocation is not null) {
            revocationStatus = CombineRevocation(revocationStatus, MapStatus(revocation.Status));
            if (revocation.TimestampAuthorityStatus == SecurityValidationStatus.Invalid &&
                timestampStatus != PdfCryptographicValidationStatus.NotPerformed) {
                timestampStatus = PdfCryptographicValidationStatus.Invalid;
            } else if (timestampStatus == PdfCryptographicValidationStatus.Indeterminate &&
                revocation.TimestampAuthorityStatus == SecurityValidationStatus.Valid &&
                security.Signers.SelectMany(static signer => signer.TimestampTokens).All(IsSettledByRevocationEvidence)) {
                // The token checks passed and only the platform revocation answer was unknown; the evidence settled it.
                timestampStatus = PdfCryptographicValidationStatus.Valid;
            }
            if (revocation.SigningCertificateMismatch) {
                // RFC 5035: a signing-certificate attribute naming another certificate invalidates the signature.
                signatureStatus = PdfCryptographicValidationStatus.Invalid;
            }
        }

        return WithPadesFacts(new PdfSignatureCryptographicResult(
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
            findings), revocation);
    }

    /// <summary>
    /// The signer path is checked at the earliest trusted signature time-stamp, the best proof of when the
    /// signature existed, unless the caller fixed a verification time. A container with several signers has
    /// no single signature time, so its signers are checked at the current time.
    /// </summary>
    private DateTimeOffset? ResolveSignerValidationTime(IReadOnlyList<CmsSignerVerificationResult> signers) {
        if (_options.CertificateValidation.VerificationTime is DateTime fixedTime) {
            return new DateTimeOffset(DateTime.SpecifyKind(fixedTime.ToUniversalTime(), DateTimeKind.Utc));
        }
        if (signers.Count != 1) return null;
        DateTimeOffset? earliest = null;
        foreach (Rfc3161TimestampVerificationResult token in signers[0].TimestampTokens) {
            if (token.Status != SecurityValidationStatus.Valid || !token.Timestamp.HasValue) continue;
            if (!earliest.HasValue || token.Timestamp.Value < earliest.Value) earliest = token.Timestamp.Value;
        }
        return earliest;
    }

    private OfficeRevocationOutcome? ValidateRevocation(
        PdfSignatureCryptographyInput input,
        byte[] encoded,
        bool isTimestampToken,
        DateTimeOffset? signerValidationTime,
        List<PdfSignatureCryptographicFinding> findings) {
        if (_securityProvider is not IOfficeRevocationValidator validator) {
            findings.Add(Finding(
                PdfDiagnosticSeverity.Info,
                "revocation-data-not-evaluated",
                _securityProvider.Name + " does not evaluate OCSP responses or CRLs from the signature or document; only its certificate policy applies."));
            return null;
        }

        DateTime? fixedTime = _options.CertificateValidation.VerificationTime;
        var request = new OfficeRevocationRequest(encoded, isTimestampToken) {
            SignerValidationTime = signerValidationTime,
            DocumentCertificates = input.ValidationData.Certificates,
            DocumentTokens = input.ValidationData.CreateTokens(),
            Policy = RevocationPolicy,
            Fetcher = RevocationFetcher,
            ExtraCertificates = _options.CertificateValidation.ExtraCertificates,
            Now = fixedTime.HasValue
                ? new DateTimeOffset(DateTime.SpecifyKind(fixedTime.Value.ToUniversalTime(), DateTimeKind.Utc))
                : null
        };
        OfficeRevocationOutcome outcome = validator.ValidateRevocation(request);
        findings.AddRange(outcome.Findings.Select(MapFinding));
        return outcome;
    }

    private static PdfSignatureCryptographicResult WithPadesFacts(
        PdfSignatureCryptographicResult result,
        OfficeRevocationOutcome? revocation) {
        if (revocation is null) return result;
        result.ReportsPadesEvidence = true;
        result.HasSigningCertificateAttribute = revocation.HasSigningCertificateAttribute;
        result.RevocationSettledByDocument = revocation.SettledByDocument;
        result.RevocationEvidenceHasProofOfExistence = revocation.DocumentEvidenceHasProofOfExistence;
        return result;
    }

    /// <summary>
    /// A time-stamp token is indeterminate with an accepted TSA chain only when the platform revocation answer was
    /// unknown (see the RFC 3161 verifier), so valid revocation evidence settles it.
    /// </summary>
    private static bool IsSettledByRevocationEvidence(Rfc3161TimestampVerificationResult token) =>
        token.Status == SecurityValidationStatus.Valid ||
        (token.Status == SecurityValidationStatus.Indeterminate &&
         token.CertificateValidation.ChainStatus == SecurityValidationStatus.Valid &&
         token.CertificateValidation.RevocationStatus is SecurityValidationStatus.Indeterminate);

    /// <summary>Embedded and fetched revocation data decide; a platform revocation check can only add a revoked result.</summary>
    private static PdfCryptographicValidationStatus CombineRevocation(
        PdfCryptographicValidationStatus platform,
        PdfCryptographicValidationStatus evidence) =>
        platform == PdfCryptographicValidationStatus.Invalid ? PdfCryptographicValidationStatus.Invalid : evidence;

    private PdfSignatureCryptographicResult VerifyDocumentTimestamp(PdfSignatureCryptographyInput input, byte[] encoded) {
        byte[] signedContent = input.SignedContent;
        if (!_options.ValidateTimestamps) {
            return new PdfSignatureCryptographicResult(
                Name,
                PdfCryptographicValidationStatus.NotPerformed,
                PdfCryptographicValidationStatus.NotPerformed,
                PdfCryptographicValidationStatus.NotPerformed,
                PdfCryptographicValidationStatus.NotPerformed,
                PdfCryptographicValidationStatus.NotPerformed,
                findings: new[] {
                    Finding(PdfDiagnosticSeverity.Info, "timestamp-validation-disabled", "RFC 3161 validation was disabled by provider policy.")
                });
        }

        Rfc3161TimestampVerificationResult timestamp = _securityProvider.VerifyTimestamp(
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
        PdfCryptographicValidationStatus revocationStatus = MapStatus(timestamp.CertificateValidation.RevocationStatus);
        OfficeRevocationOutcome? revocation = ValidateRevocation(
            input,
            encoded,
            isTimestampToken: true,
            signerValidationTime: null,
            findings);
        if (revocation is not null) {
            revocationStatus = CombineRevocation(revocationStatus, MapStatus(revocation.Status));
            if (revocationStatus == PdfCryptographicValidationStatus.Invalid) {
                timestampStatus = PdfCryptographicValidationStatus.Invalid;
            } else if (revocationStatus == PdfCryptographicValidationStatus.Valid && IsSettledByRevocationEvidence(timestamp)) {
                // Signature, imprint and TSA chain passed; only the platform revocation answer was unknown.
                cryptoStatus = PdfCryptographicValidationStatus.Valid;
                timestampStatus = PdfCryptographicValidationStatus.Valid;
            }
        }
        if (timestampStatus == PdfCryptographicValidationStatus.Invalid) {
            findings.Add(Finding(
                PdfDiagnosticSeverity.Error,
                "timestamp-message-imprint-invalid",
                "The RFC 3161 token, TSA trust policy, or message imprint did not validate against the PDF signed byte ranges."));
        }

        return WithPadesFacts(new PdfSignatureCryptographicResult(
            Name,
            cryptoStatus,
            cryptoStatus,
            chainStatus,
            revocationStatus,
            timestampStatus,
            timestampTime: timestamp.Timestamp,
            findings: findings), revocation);
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
        if (result.Any(static finding => finding.Code == "signer-certificate-missing")) {
            result.Add(Finding(PdfDiagnosticSeverity.Error, "cms-signer-missing", "The CMS signer certificate was not embedded or supplied by caller policy."));
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
