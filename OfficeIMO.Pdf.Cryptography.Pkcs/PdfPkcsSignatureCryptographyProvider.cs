using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.IO;

namespace OfficeIMO.Pdf.Cryptography;

#pragma warning disable CA1510 // Cross-target guard code supports netstandard2.0 and net472.

/// <summary>First-party managed CMS, RFC 3161, and X509Chain implementation of the OfficeIMO.Pdf cryptography seam.</summary>
public sealed class PdfPkcsSignatureCryptographyProvider : IPdfSignatureCryptographyProvider {
    private const string Sha1Oid = "1.3.14.3.2.26";
    private const string Sha256Oid = "2.16.840.1.101.3.4.2.1";
    private const string Sha384Oid = "2.16.840.1.101.3.4.2.2";
    private const string Sha512Oid = "2.16.840.1.101.3.4.2.3";
    private const string RsaEncryptionOid = "1.2.840.113549.1.1.1";
    private const string Sha1WithRsaOid = "1.2.840.113549.1.1.5";
    private const string Sha256WithRsaOid = "1.2.840.113549.1.1.11";
    private const string Sha384WithRsaOid = "1.2.840.113549.1.1.12";
    private const string Sha512WithRsaOid = "1.2.840.113549.1.1.13";
    private readonly PdfPkcsSignatureValidationOptions _options;

    /// <summary>Creates a provider using caller policy or conservative no-network defaults.</summary>
    public PdfPkcsSignatureCryptographyProvider(PdfPkcsSignatureValidationOptions? options = null) {
        _options = options ?? new PdfPkcsSignatureValidationOptions();
    }

    /// <inheritdoc />
    public string Name => "OfficeIMO.Pdf managed PKCS";

    /// <inheritdoc />
    public PdfSignatureCryptographicResult Verify(PdfSignatureCryptographyInput input) {
        if (input == null) throw new ArgumentNullException(nameof(input));
        byte[] encodedSignature = TrimDerContainer(input.SignatureContents);
        return input.Signature.IsDocumentTimestamp
            ? VerifyDocumentTimestamp(input, encodedSignature)
            : VerifyCmsSignature(input, encodedSignature);
    }

    private PdfSignatureCryptographicResult VerifyCmsSignature(PdfSignatureCryptographyInput input, byte[] encodedSignature) {
        var findings = new List<PdfSignatureCryptographicFinding>();
        PdfManagedCmsDocument cms;
        try {
            cms = PdfManagedCmsDocument.Parse(encodedSignature);
        } catch (Exception ex) when (ex is InvalidDataException || ex is CryptographicException || ex is ArgumentException) {
            findings.Add(Finding(PdfDiagnosticSeverity.Error, "CmsSignatureInvalid", "CMS signature container could not be decoded: " + ex.Message));
            return InvalidResult(findings);
        }

        bool encapsulatedSha1 = string.Equals(input.Signature.SubFilter, "adbe.pkcs7.sha1", StringComparison.Ordinal);
        byte[] cmsContent = cms.EncapsulatedContent ?? input.SignedContent;
        PdfCryptographicValidationStatus digestStatus = VerifyMessageDigest(cms, cmsContent);
        if (encapsulatedSha1) {
            digestStatus = cms.EncapsulatedContent != null && FixedTimeEquals(Hash(input.SignedContent, HashAlgorithmName.SHA1), cms.EncapsulatedContent)
                ? digestStatus
                : PdfCryptographicValidationStatus.Invalid;
        }
        if (digestStatus == PdfCryptographicValidationStatus.Invalid) {
            findings.Add(Finding(PdfDiagnosticSeverity.Error, "CmsDigestMismatch", "CMS message digest does not match the PDF signed byte ranges."));
        }

        X509Certificate2? certificate = cms.FindSignerCertificate(_options.ExtraCertificates);
        PdfCryptographicValidationStatus mathStatus = VerifyMathematicalSignature(cms, cmsContent, certificate, findings);
        if (digestStatus == PdfCryptographicValidationStatus.Invalid) {
            mathStatus = PdfCryptographicValidationStatus.Invalid;
            findings.Add(Finding(PdfDiagnosticSeverity.Error, "CmsSignatureInvalid", "CMS signed-attribute validation failed because the content digest does not match."));
        }
        ChainResult chain = ValidateCertificate(certificate, findings);
        TimestampResult timestamp = _options.ValidateTimestamps
            ? ValidateSignatureTimestamp(cms, findings)
            : TimestampResult.NotPerformed;
        return CreateResult(mathStatus, digestStatus, chain, timestamp, certificate, cms.SigningTime, findings);
    }

    private PdfSignatureCryptographicResult VerifyDocumentTimestamp(PdfSignatureCryptographyInput input, byte[] encodedSignature) {
        var findings = new List<PdfSignatureCryptographicFinding>();
        if (!_options.ValidateTimestamps) {
            findings.Add(Finding(PdfDiagnosticSeverity.Info, "TimestampValidationDisabled", "RFC 3161 validation was disabled by provider policy."));
            return new PdfSignatureCryptographicResult(Name, PdfCryptographicValidationStatus.NotPerformed, PdfCryptographicValidationStatus.NotPerformed, PdfCryptographicValidationStatus.NotPerformed, PdfCryptographicValidationStatus.NotPerformed, PdfCryptographicValidationStatus.NotPerformed, findings: findings.AsReadOnly());
        }

        if (!TryVerifyTimestampToken(encodedSignature, input.SignedContent, findings, out X509Certificate2? certificate, out DateTimeOffset? timestampTime)) {
            findings.Add(Finding(PdfDiagnosticSeverity.Error, "TimestampMessageImprintInvalid", "RFC 3161 signature, TSA certificate, or message imprint did not validate against the PDF signed byte ranges."));
            return InvalidTimestampResult(findings);
        }

        ChainResult chain = ValidateCertificate(certificate, findings);
        return CreateResult(PdfCryptographicValidationStatus.Valid, PdfCryptographicValidationStatus.Valid, chain, new TimestampResult(PdfCryptographicValidationStatus.Valid, timestampTime), certificate, null, findings);
    }

    private static PdfCryptographicValidationStatus VerifyMessageDigest(PdfManagedCmsDocument cms, byte[] content) {
        if (cms.MessageDigest == null) return PdfCryptographicValidationStatus.Valid;
        if (!TryGetHashAlgorithm(cms.DigestAlgorithmOid, out HashAlgorithmName algorithm)) return PdfCryptographicValidationStatus.Indeterminate;
        return FixedTimeEquals(Hash(content, algorithm), cms.MessageDigest)
            ? PdfCryptographicValidationStatus.Valid
            : PdfCryptographicValidationStatus.Invalid;
    }

    private static PdfCryptographicValidationStatus VerifyMathematicalSignature(
        PdfManagedCmsDocument cms,
        byte[] content,
        X509Certificate2? certificate,
        List<PdfSignatureCryptographicFinding> findings) {
        if (certificate == null) {
            findings.Add(Finding(PdfDiagnosticSeverity.Error, "CmsSignerMissing", "CMS signer certificate was not embedded or supplied by caller policy."));
            return PdfCryptographicValidationStatus.Invalid;
        }
        if (!TryGetHashAlgorithm(cms.DigestAlgorithmOid, out HashAlgorithmName algorithm) || !IsRsaSignatureAlgorithm(cms.SignatureAlgorithmOid)) {
            findings.Add(Finding(PdfDiagnosticSeverity.Warning, "CmsAlgorithmUnsupported", "CMS digest or signature algorithm is not supported by the managed provider."));
            return PdfCryptographicValidationStatus.Indeterminate;
        }
        using RSA? rsa = certificate.GetRSAPublicKey();
        if (rsa == null) {
            findings.Add(Finding(PdfDiagnosticSeverity.Warning, "CmsAlgorithmUnsupported", "CMS signer certificate does not expose an RSA public key."));
            return PdfCryptographicValidationStatus.Indeterminate;
        }
        bool valid;
        try {
            valid = rsa.VerifyData(cms.SignedAttributes ?? content, cms.SignatureValue, algorithm, RSASignaturePadding.Pkcs1);
        } catch (CryptographicException) {
            valid = false;
        }
        if (!valid) findings.Add(Finding(PdfDiagnosticSeverity.Error, "CmsSignatureInvalid", "CMS mathematical signature validation failed."));
        return valid ? PdfCryptographicValidationStatus.Valid : PdfCryptographicValidationStatus.Invalid;
    }

    private TimestampResult ValidateSignatureTimestamp(PdfManagedCmsDocument cms, List<PdfSignatureCryptographicFinding> findings) {
        if (cms.SignatureTimestamps.Count == 0) return TimestampResult.NotPerformed;
        for (int i = 0; i < cms.SignatureTimestamps.Count; i++) {
            if (TryVerifyTimestampToken(cms.SignatureTimestamps[i], cms.SignatureValue, findings, out _, out DateTimeOffset? timestamp)) {
                return new TimestampResult(PdfCryptographicValidationStatus.Valid, timestamp);
            }
        }
        findings.Add(Finding(PdfDiagnosticSeverity.Error, "SignatureTimestampInvalid", "The CMS signature-timestamp attribute did not validate against the signer signature value."));
        return new TimestampResult(PdfCryptographicValidationStatus.Invalid, null);
    }

    private bool TryVerifyTimestampToken(
        byte[] encoded,
        byte[] expectedData,
        List<PdfSignatureCryptographicFinding> findings,
        out X509Certificate2? certificate,
        out DateTimeOffset? timestamp) {
        certificate = null;
        timestamp = null;
        try {
            PdfManagedCmsDocument cms = PdfManagedCmsDocument.Parse(encoded);
            if (!string.Equals(cms.ContentTypeOid, PdfManagedCmsDocument.TstInfoOid, StringComparison.Ordinal) || cms.EncapsulatedContent == null) return false;
            certificate = cms.FindSignerCertificate(_options.ExtraCertificates);
            if (VerifyMessageDigest(cms, cms.EncapsulatedContent) != PdfCryptographicValidationStatus.Valid ||
                VerifyMathematicalSignature(cms, cms.EncapsulatedContent, certificate, findings) != PdfCryptographicValidationStatus.Valid) return false;
            return TryReadTimestampInfo(cms.EncapsulatedContent, expectedData, out timestamp);
        } catch (Exception ex) when (ex is InvalidDataException || ex is CryptographicException || ex is ArgumentException) {
            return false;
        }
    }

    private static bool TryReadTimestampInfo(byte[] encoded, byte[] expectedData, out DateTimeOffset? timestamp) {
        timestamp = null;
        var root = new PdfDerReader(encoded);
        var info = root.Read(0x30).Reader();
        info.Read(0x02);
        info.Read(0x06);
        var imprint = info.Read(0x30).Reader();
        string algorithmOid = PdfManagedCmsDocument.ReadAlgorithmIdentifier(imprint.Read(0x30));
        byte[] actualDigest = imprint.Read(0x04).Content();
        imprint.EnsureEnd();
        info.Read(0x02);
        PdfDerElement time = info.Read();
        timestamp = PdfManagedCmsDocument.ReadTime(time);
        return timestamp.HasValue && TryGetHashAlgorithm(algorithmOid, out HashAlgorithmName algorithm) && FixedTimeEquals(Hash(expectedData, algorithm), actualDigest);
    }

    private ChainResult ValidateCertificate(X509Certificate2? certificate, List<PdfSignatureCryptographicFinding> findings) {
        if (certificate == null) return new ChainResult(PdfCryptographicValidationStatus.Indeterminate, PdfCryptographicValidationStatus.NotPerformed);
        if (!_options.ValidateCertificateChain) return new ChainResult(PdfCryptographicValidationStatus.NotPerformed, PdfCryptographicValidationStatus.NotPerformed);
        using var chain = new X509Chain();
        chain.ChainPolicy.RevocationMode = _options.RevocationMode;
        chain.ChainPolicy.RevocationFlag = _options.RevocationFlag;
        chain.ChainPolicy.VerificationFlags = _options.VerificationFlags;
        chain.ChainPolicy.UrlRetrievalTimeout = _options.UrlRetrievalTimeout;
        if (_options.VerificationTime.HasValue) chain.ChainPolicy.VerificationTime = _options.VerificationTime.Value;
        chain.ChainPolicy.ExtraStore.AddRange(_options.ExtraCertificates);
        bool platformResult = chain.Build(certificate);
        bool accepted = _options.ChainEvaluator?.Invoke(certificate, chain) ?? platformResult;
        if (!accepted) {
            string statuses = chain.ChainStatus.Length == 0 ? "no platform chain status" : string.Join(", ", chain.ChainStatus.Select(static status => status.Status.ToString()));
            findings.Add(Finding(PdfDiagnosticSeverity.Warning, "CertificateChainUntrusted", "Signer certificate chain was not accepted: " + statuses + "."));
        }
        return new ChainResult(accepted ? PdfCryptographicValidationStatus.Valid : PdfCryptographicValidationStatus.Invalid, ClassifyRevocation(chain));
    }

    private PdfCryptographicValidationStatus ClassifyRevocation(X509Chain chain) {
        if (_options.RevocationMode == X509RevocationMode.NoCheck) return PdfCryptographicValidationStatus.NotPerformed;
        bool indeterminate = false;
        foreach (X509ChainStatus status in chain.ChainStatus) {
            if ((status.Status & X509ChainStatusFlags.Revoked) != 0) return PdfCryptographicValidationStatus.Invalid;
            if ((status.Status & (X509ChainStatusFlags.RevocationStatusUnknown | X509ChainStatusFlags.OfflineRevocation)) != 0) indeterminate = true;
        }
        return indeterminate ? PdfCryptographicValidationStatus.Indeterminate : PdfCryptographicValidationStatus.Valid;
    }

    private PdfSignatureCryptographicResult CreateResult(PdfCryptographicValidationStatus mathStatus, PdfCryptographicValidationStatus digestStatus, ChainResult chain, TimestampResult timestamp, X509Certificate2? certificate, DateTimeOffset? signingTime, List<PdfSignatureCryptographicFinding> findings) =>
        new PdfSignatureCryptographicResult(Name, mathStatus, digestStatus, chain.ChainStatus, chain.RevocationStatus, timestamp.Status, certificate?.Subject, certificate?.Issuer, certificate?.SerialNumber, certificate?.Thumbprint, signingTime, timestamp.Time, findings.AsReadOnly());
    private PdfSignatureCryptographicResult InvalidResult(List<PdfSignatureCryptographicFinding> findings) =>
        new PdfSignatureCryptographicResult(Name, PdfCryptographicValidationStatus.Invalid, PdfCryptographicValidationStatus.Invalid, PdfCryptographicValidationStatus.NotPerformed, PdfCryptographicValidationStatus.NotPerformed, PdfCryptographicValidationStatus.NotPerformed, findings: findings.AsReadOnly());
    private PdfSignatureCryptographicResult InvalidTimestampResult(List<PdfSignatureCryptographicFinding> findings) =>
        new PdfSignatureCryptographicResult(Name, PdfCryptographicValidationStatus.Invalid, PdfCryptographicValidationStatus.Invalid, PdfCryptographicValidationStatus.NotPerformed, PdfCryptographicValidationStatus.NotPerformed, PdfCryptographicValidationStatus.Invalid, findings: findings.AsReadOnly());

    private static bool TryGetHashAlgorithm(string oid, out HashAlgorithmName algorithm) {
        switch (oid) {
            case Sha1Oid: algorithm = HashAlgorithmName.SHA1; return true;
            case Sha256Oid: algorithm = HashAlgorithmName.SHA256; return true;
            case Sha384Oid: algorithm = HashAlgorithmName.SHA384; return true;
            case Sha512Oid: algorithm = HashAlgorithmName.SHA512; return true;
            default: algorithm = default; return false;
        }
    }
    private static bool IsRsaSignatureAlgorithm(string oid) => oid == RsaEncryptionOid || oid == Sha1WithRsaOid || oid == Sha256WithRsaOid || oid == Sha384WithRsaOid || oid == Sha512WithRsaOid;
    #pragma warning disable CA5350 // SHA-1 is required only to validate legacy authored PDF signatures.
    private static byte[] Hash(byte[] data, HashAlgorithmName algorithm) {
        using HashAlgorithm hash = algorithm == HashAlgorithmName.SHA1 ? SHA1.Create() : algorithm == HashAlgorithmName.SHA256 ? SHA256.Create() : algorithm == HashAlgorithmName.SHA384 ? SHA384.Create() : algorithm == HashAlgorithmName.SHA512 ? SHA512.Create() : throw new NotSupportedException("Hash algorithm is not supported.");
        return hash.ComputeHash(data);
    }
    #pragma warning restore CA5350
    private static bool FixedTimeEquals(byte[] left, byte[] right) => PdfManagedCmsDocument.FixedTimeEquals(left, right);
    private static byte[] TrimDerContainer(byte[] value) {
        try { return new PdfDerReader(value).Read(0x30).Encoded(); } catch (InvalidDataException) { return (byte[])value.Clone(); }
    }
    private static PdfSignatureCryptographicFinding Finding(PdfDiagnosticSeverity severity, string code, string message) => new PdfSignatureCryptographicFinding(severity, code, message);

    private sealed class ChainResult {
        internal ChainResult(PdfCryptographicValidationStatus chainStatus, PdfCryptographicValidationStatus revocationStatus) { ChainStatus = chainStatus; RevocationStatus = revocationStatus; }
        internal PdfCryptographicValidationStatus ChainStatus { get; }
        internal PdfCryptographicValidationStatus RevocationStatus { get; }
    }
    private sealed class TimestampResult {
        internal static readonly TimestampResult NotPerformed = new TimestampResult(PdfCryptographicValidationStatus.NotPerformed, null);
        internal TimestampResult(PdfCryptographicValidationStatus status, DateTimeOffset? time) { Status = status; Time = time; }
        internal PdfCryptographicValidationStatus Status { get; }
        internal DateTimeOffset? Time { get; }
    }
}
