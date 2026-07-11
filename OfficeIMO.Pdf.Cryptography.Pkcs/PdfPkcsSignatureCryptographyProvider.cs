using System.Security.Cryptography;
using System.Security.Cryptography.Pkcs;
using System.Security.Cryptography.X509Certificates;

namespace OfficeIMO.Pdf.Cryptography;

/// <summary>Optional SignedCms, RFC 3161, and X509Chain implementation of the OfficeIMO.Pdf cryptography seam.</summary>
public sealed class PdfPkcsSignatureCryptographyProvider : IPdfSignatureCryptographyProvider {
    private const string SigningTimeOid = "1.2.840.113549.1.9.5";
    private const string SignatureTimestampOid = "1.2.840.113549.1.9.16.2.14";
    private readonly PdfPkcsSignatureValidationOptions _options;

    /// <summary>Creates a provider using caller policy or conservative no-network defaults.</summary>
    public PdfPkcsSignatureCryptographyProvider(PdfPkcsSignatureValidationOptions? options = null) {
        _options = options ?? new PdfPkcsSignatureValidationOptions();
    }

    /// <inheritdoc />
    public string Name => "System.Security.Cryptography.Pkcs";

    /// <inheritdoc />
    public PdfSignatureCryptographicResult Verify(PdfSignatureCryptographyInput input) {
#if NET8_0_OR_GREATER
        ArgumentNullException.ThrowIfNull(input);
#else
        if (input is null) throw new ArgumentNullException(nameof(input));
#endif

        byte[] encodedSignature = TrimDerContainer(input.SignatureContents);
        return input.Signature.IsDocumentTimestamp
            ? VerifyDocumentTimestamp(input, encodedSignature)
            : VerifyCmsSignature(input, encodedSignature);
    }

    private PdfSignatureCryptographicResult VerifyCmsSignature(
        PdfSignatureCryptographyInput input,
        byte[] encodedSignature) {
        var findings = new List<PdfSignatureCryptographicFinding>();
        PdfCryptographicValidationStatus mathStatus;
        PdfCryptographicValidationStatus digestStatus;
        SignedCms cms;
        try {
            bool encapsulatedSha1 = string.Equals(input.Signature.SubFilter, "adbe.pkcs7.sha1", StringComparison.Ordinal);
            cms = encapsulatedSha1
                ? new SignedCms()
                : new SignedCms(new ContentInfo(input.SignedContent), detached: true);
            cms.Decode(encodedSignature);
            cms.CheckSignature(_options.ExtraCertificates, verifySignatureOnly: true);
            mathStatus = PdfCryptographicValidationStatus.Valid;
            digestStatus = encapsulatedSha1
                ? VerifySha1Content(cms.ContentInfo.Content, input.SignedContent)
                : PdfCryptographicValidationStatus.Valid;
            if (digestStatus != PdfCryptographicValidationStatus.Valid) {
                findings.Add(Finding(PdfDiagnosticSeverity.Error, "CmsDigestMismatch", "CMS encapsulated SHA-1 content does not match the PDF signed byte ranges."));
            }
        } catch (CryptographicException ex) {
            findings.Add(Finding(PdfDiagnosticSeverity.Error, "CmsSignatureInvalid", "CMS signature or signed-attribute validation failed: " + ex.Message));
            return InvalidResult(findings);
        }

        if (cms.SignerInfos.Count == 0) {
            findings.Add(Finding(PdfDiagnosticSeverity.Error, "CmsSignerMissing", "CMS content does not contain a signer."));
            return InvalidResult(findings);
        }

        SignerInfo signer = cms.SignerInfos[0];
        X509Certificate2? certificate = signer.Certificate;
        ChainResult chain = ValidateCertificate(certificate, findings);
        DateTimeOffset? signingTime = ReadSigningTime(signer);
        TimestampResult timestamp = _options.ValidateTimestamps
            ? ValidateSignatureTimestamp(signer, findings)
            : TimestampResult.NotPerformed;

        return CreateResult(
            mathStatus,
            digestStatus,
            chain,
            timestamp,
            certificate,
            signingTime,
            findings);
    }

    private PdfSignatureCryptographicResult VerifyDocumentTimestamp(
        PdfSignatureCryptographyInput input,
        byte[] encodedSignature) {
        var findings = new List<PdfSignatureCryptographicFinding>();
#if !NET8_0_OR_GREATER
        findings.Add(Finding(PdfDiagnosticSeverity.Warning, "TimestampApiUnavailable", "RFC 3161 validation requires the optional provider's .NET 8 or later target."));
        return new PdfSignatureCryptographicResult(
            Name,
            PdfCryptographicValidationStatus.NotPerformed,
            PdfCryptographicValidationStatus.NotPerformed,
            PdfCryptographicValidationStatus.NotPerformed,
            PdfCryptographicValidationStatus.NotPerformed,
            PdfCryptographicValidationStatus.Indeterminate,
            findings: findings.AsReadOnly());
#else
        if (!_options.ValidateTimestamps) {
            findings.Add(Finding(PdfDiagnosticSeverity.Info, "TimestampValidationDisabled", "RFC 3161 validation was disabled by provider policy."));
            return new PdfSignatureCryptographicResult(
                Name,
                PdfCryptographicValidationStatus.NotPerformed,
                PdfCryptographicValidationStatus.NotPerformed,
                PdfCryptographicValidationStatus.NotPerformed,
                PdfCryptographicValidationStatus.NotPerformed,
                PdfCryptographicValidationStatus.NotPerformed,
                findings: findings.AsReadOnly());
        }

        if (!Rfc3161TimestampToken.TryDecode(encodedSignature, out Rfc3161TimestampToken? token, out int consumed) ||
            token is null ||
            consumed != encodedSignature.Length) {
            findings.Add(Finding(PdfDiagnosticSeverity.Error, "TimestampTokenInvalid", "Signature /Contents is not one complete DER-encoded RFC 3161 timestamp token."));
            return InvalidTimestampResult(findings);
        }

        bool valid = token.VerifySignatureForData(input.SignedContent, out X509Certificate2? certificate, _options.ExtraCertificates);
        if (!valid || certificate is null) {
            findings.Add(Finding(PdfDiagnosticSeverity.Error, "TimestampMessageImprintInvalid", "RFC 3161 signature, TSA certificate, or message imprint did not validate against the PDF signed byte ranges."));
            return InvalidTimestampResult(findings);
        }

        ChainResult chain = ValidateCertificate(certificate, findings);
        TimestampResult timestamp = new TimestampResult(PdfCryptographicValidationStatus.Valid, token.TokenInfo.Timestamp);
        return CreateResult(
            PdfCryptographicValidationStatus.Valid,
            PdfCryptographicValidationStatus.Valid,
            chain,
            timestamp,
            certificate,
            signingTime: null,
            findings);
#endif
    }

#if NET8_0_OR_GREATER
    private TimestampResult ValidateSignatureTimestamp(
#else
    private static TimestampResult ValidateSignatureTimestamp(
#endif
        SignerInfo signer,
        List<PdfSignatureCryptographicFinding> findings) {
#if !NET8_0_OR_GREATER
        foreach (CryptographicAttributeObject attribute in signer.UnsignedAttributes) {
            if (string.Equals(attribute.Oid?.Value, SignatureTimestampOid, StringComparison.Ordinal)) {
                findings.Add(Finding(PdfDiagnosticSeverity.Warning, "TimestampApiUnavailable", "CMS signature-timestamp validation requires the optional provider's .NET 8 or later target."));
                return new TimestampResult(PdfCryptographicValidationStatus.Indeterminate, null);
            }
        }

        return TimestampResult.NotPerformed;
#else
        foreach (CryptographicAttributeObject attribute in signer.UnsignedAttributes) {
            if (!string.Equals(attribute.Oid?.Value, SignatureTimestampOid, StringComparison.Ordinal)) {
                continue;
            }

            foreach (AsnEncodedData value in attribute.Values) {
                if (!Rfc3161TimestampToken.TryDecode(value.RawData, out Rfc3161TimestampToken? token, out int consumed) ||
                    token is null ||
                    consumed != value.RawData.Length) {
                    continue;
                }

                bool valid = token.VerifySignatureForSignerInfo(signer, out X509Certificate2? _, _options.ExtraCertificates);
                if (valid) {
                    return new TimestampResult(PdfCryptographicValidationStatus.Valid, token.TokenInfo.Timestamp);
                }
            }

            findings.Add(Finding(PdfDiagnosticSeverity.Error, "SignatureTimestampInvalid", "The CMS signature-timestamp attribute did not validate against the signer signature value."));
            return new TimestampResult(PdfCryptographicValidationStatus.Invalid, null);
        }

        return TimestampResult.NotPerformed;
#endif
    }

    private ChainResult ValidateCertificate(
        X509Certificate2? certificate,
        List<PdfSignatureCryptographicFinding> findings) {
        if (certificate is null) {
            findings.Add(Finding(PdfDiagnosticSeverity.Warning, "SignerCertificateMissing", "The CMS signer certificate was not embedded or supplied by caller policy."));
            return new ChainResult(PdfCryptographicValidationStatus.Indeterminate, PdfCryptographicValidationStatus.NotPerformed);
        }

        if (!_options.ValidateCertificateChain) {
            return new ChainResult(PdfCryptographicValidationStatus.NotPerformed, PdfCryptographicValidationStatus.NotPerformed);
        }

        using var chain = new X509Chain();
        chain.ChainPolicy.RevocationMode = _options.RevocationMode;
        chain.ChainPolicy.RevocationFlag = _options.RevocationFlag;
        chain.ChainPolicy.VerificationFlags = _options.VerificationFlags;
        chain.ChainPolicy.UrlRetrievalTimeout = _options.UrlRetrievalTimeout;
        if (_options.VerificationTime.HasValue) {
            chain.ChainPolicy.VerificationTime = _options.VerificationTime.Value;
        }

        chain.ChainPolicy.ExtraStore.AddRange(_options.ExtraCertificates);
        bool platformResult = chain.Build(certificate);
        bool accepted = _options.ChainEvaluator?.Invoke(certificate, chain) ?? platformResult;
        PdfCryptographicValidationStatus chainStatus = accepted
            ? PdfCryptographicValidationStatus.Valid
            : PdfCryptographicValidationStatus.Invalid;
        PdfCryptographicValidationStatus revocationStatus = ClassifyRevocation(chain);
        if (!accepted) {
            string statuses = chain.ChainStatus.Length == 0
                ? "no platform chain status"
                : string.Join(", ", chain.ChainStatus.Select(static status => status.Status.ToString()));
            findings.Add(Finding(PdfDiagnosticSeverity.Warning, "CertificateChainUntrusted", "Signer certificate chain was not accepted: " + statuses + "."));
        }

        return new ChainResult(chainStatus, revocationStatus);
    }

    private PdfCryptographicValidationStatus ClassifyRevocation(X509Chain chain) {
        if (_options.RevocationMode == X509RevocationMode.NoCheck) {
            return PdfCryptographicValidationStatus.NotPerformed;
        }

        bool indeterminate = false;
        foreach (X509ChainStatus status in chain.ChainStatus) {
            if ((status.Status & X509ChainStatusFlags.Revoked) != 0) {
                return PdfCryptographicValidationStatus.Invalid;
            }

            if ((status.Status & (X509ChainStatusFlags.RevocationStatusUnknown | X509ChainStatusFlags.OfflineRevocation)) != 0) {
                indeterminate = true;
            }
        }

        return indeterminate
            ? PdfCryptographicValidationStatus.Indeterminate
            : PdfCryptographicValidationStatus.Valid;
    }

    private PdfSignatureCryptographicResult CreateResult(
        PdfCryptographicValidationStatus mathStatus,
        PdfCryptographicValidationStatus digestStatus,
        ChainResult chain,
        TimestampResult timestamp,
        X509Certificate2? certificate,
        DateTimeOffset? signingTime,
        List<PdfSignatureCryptographicFinding> findings) {
        return new PdfSignatureCryptographicResult(
            Name,
            mathStatus,
            digestStatus,
            chain.ChainStatus,
            chain.RevocationStatus,
            timestamp.Status,
            certificate?.Subject,
            certificate?.Issuer,
            certificate?.SerialNumber,
            certificate?.Thumbprint,
            signingTime,
            timestamp.Time,
            findings.AsReadOnly());
    }

    private PdfSignatureCryptographicResult InvalidResult(List<PdfSignatureCryptographicFinding> findings) {
        return new PdfSignatureCryptographicResult(
            Name,
            PdfCryptographicValidationStatus.Invalid,
            PdfCryptographicValidationStatus.Invalid,
            PdfCryptographicValidationStatus.NotPerformed,
            PdfCryptographicValidationStatus.NotPerformed,
            PdfCryptographicValidationStatus.NotPerformed,
            findings: findings.AsReadOnly());
    }

    private PdfSignatureCryptographicResult InvalidTimestampResult(List<PdfSignatureCryptographicFinding> findings) {
        return new PdfSignatureCryptographicResult(
            Name,
            PdfCryptographicValidationStatus.Invalid,
            PdfCryptographicValidationStatus.Invalid,
            PdfCryptographicValidationStatus.NotPerformed,
            PdfCryptographicValidationStatus.NotPerformed,
            PdfCryptographicValidationStatus.Invalid,
            findings: findings.AsReadOnly());
    }

    private static DateTimeOffset? ReadSigningTime(SignerInfo signer) {
        foreach (CryptographicAttributeObject attribute in signer.SignedAttributes) {
            if (!string.Equals(attribute.Oid?.Value, SigningTimeOid, StringComparison.Ordinal) || attribute.Values.Count == 0) {
                continue;
            }

            try {
                return new DateTimeOffset(new Pkcs9SigningTime(attribute.Values[0].RawData).SigningTime);
            } catch (CryptographicException) {
                return null;
            }
        }

        return null;
    }

    private static PdfCryptographicValidationStatus VerifySha1Content(byte[] encapsulatedDigest, byte[] signedContent) {
#pragma warning disable CA5350, CA1850
        using SHA1 sha1 = SHA1.Create();
        byte[] expected = sha1.ComputeHash(signedContent);
#pragma warning restore CA5350, CA1850
        return FixedTimeEquals(expected, encapsulatedDigest)
            ? PdfCryptographicValidationStatus.Valid
            : PdfCryptographicValidationStatus.Invalid;
    }

    private static bool FixedTimeEquals(byte[] left, byte[] right) {
        if (left.Length != right.Length) return false;
        int difference = 0;
        for (int i = 0; i < left.Length; i++) difference |= left[i] ^ right[i];
        return difference == 0;
    }

    private static byte[] TrimDerContainer(byte[] value) {
        if (value.Length < 2 || value[0] != 0x30) {
            return (byte[])value.Clone();
        }

        int offset = 1;
        int firstLength = value[offset++];
        long contentLength;
        if ((firstLength & 0x80) == 0) {
            contentLength = firstLength;
        } else {
            int lengthBytes = firstLength & 0x7F;
            if (lengthBytes == 0 || lengthBytes > 4 || offset + lengthBytes > value.Length) {
                return (byte[])value.Clone();
            }

            contentLength = 0;
            for (int i = 0; i < lengthBytes; i++) contentLength = (contentLength << 8) | value[offset++];
        }

        long totalLength = offset + contentLength;
        if (totalLength <= 0 || totalLength > value.Length || totalLength > int.MaxValue) {
            return (byte[])value.Clone();
        }

        var trimmed = new byte[(int)totalLength];
        Buffer.BlockCopy(value, 0, trimmed, 0, trimmed.Length);
        return trimmed;
    }

    private static PdfSignatureCryptographicFinding Finding(
        PdfDiagnosticSeverity severity,
        string code,
        string message) => new PdfSignatureCryptographicFinding(severity, code, message);

    private sealed class ChainResult {
        public ChainResult(PdfCryptographicValidationStatus chainStatus, PdfCryptographicValidationStatus revocationStatus) {
            ChainStatus = chainStatus;
            RevocationStatus = revocationStatus;
        }

        public PdfCryptographicValidationStatus ChainStatus { get; }
        public PdfCryptographicValidationStatus RevocationStatus { get; }
    }

    private sealed class TimestampResult {
        public static readonly TimestampResult NotPerformed = new TimestampResult(PdfCryptographicValidationStatus.NotPerformed, null);

        public TimestampResult(PdfCryptographicValidationStatus status, DateTimeOffset? time) {
            Status = status;
            Time = time;
        }

        public PdfCryptographicValidationStatus Status { get; }
        public DateTimeOffset? Time { get; }
    }
}
