using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;

namespace OfficeIMO.Security;

/// <summary>Result state used by package-level signature inspection and validation.</summary>
public enum OfficePackageSignatureValidationState {
    /// <summary>The package does not contain evidence for this check.</summary>
    NotPresent,
    /// <summary>The check was intentionally not performed.</summary>
    NotChecked,
    /// <summary>The check completed successfully.</summary>
    Passed,
    /// <summary>The check completed and rejected the package evidence.</summary>
    Failed,
    /// <summary>The package uses a profile that the current provider cannot validate.</summary>
    Unsupported
}

/// <summary>Bounded structural OPC signature inspection policy.</summary>
public sealed class OfficePackageSignatureInspectionOptions {
    /// <summary>Maximum encoded package bytes. Defaults to 512 MiB.</summary>
    public long MaxPackageBytes { get; set; } = 512L * 1024L * 1024L;
    /// <summary>Maximum package entries. Defaults to 10,000.</summary>
    public int MaxPackageParts { get; set; } = 10000;
    /// <summary>Maximum bytes read from one package part. Defaults to 256 MiB.</summary>
    public long MaxPartBytes { get; set; } = 256L * 1024L * 1024L;
    /// <summary>Maximum aggregate bytes hashed across package references. Defaults to 512 MiB.</summary>
    public long MaxTotalDigestBytes { get; set; } = 512L * 1024L * 1024L;
    /// <summary>Maximum XML signature parts. Defaults to 32.</summary>
    public int MaxSignatureParts { get; set; } = 32;
    /// <summary>Maximum bytes read from one signature part. Defaults to 16 MiB.</summary>
    public long MaxSignatureBytes { get; set; } = 16L * 1024L * 1024L;
    /// <summary>Maximum manifest references per signature. Defaults to 4,096.</summary>
    public int MaxSignedReferences { get; set; } = 4096;
    /// <summary>Maximum embedded certificates per signature. Defaults to 64.</summary>
    public int MaxCertificates { get; set; } = 64;
    /// <summary>Maximum encoded bytes for one certificate. Defaults to 4 MiB.</summary>
    public long MaxCertificateBytes { get; set; } = 4L * 1024L * 1024L;
    /// <summary>Maximum aggregate decoded certificate bytes per signature. Defaults to 64 MiB.</summary>
    public long MaxTotalCertificateBytes { get; set; } = 64L * 1024L * 1024L;
    /// <summary>Whether supported OPC reference digests are verified. Defaults to true.</summary>
    public bool VerifyDigests { get; set; } = true;

    internal void Validate() {
        if (MaxPackageBytes <= 0) throw new ArgumentOutOfRangeException(nameof(MaxPackageBytes));
        if (MaxPackageParts <= 0) throw new ArgumentOutOfRangeException(nameof(MaxPackageParts));
        if (MaxPartBytes <= 0) throw new ArgumentOutOfRangeException(nameof(MaxPartBytes));
        if (MaxTotalDigestBytes <= 0) throw new ArgumentOutOfRangeException(nameof(MaxTotalDigestBytes));
        if (MaxSignatureParts <= 0) throw new ArgumentOutOfRangeException(nameof(MaxSignatureParts));
        if (MaxSignatureBytes <= 0) throw new ArgumentOutOfRangeException(nameof(MaxSignatureBytes));
        if (MaxSignedReferences <= 0) throw new ArgumentOutOfRangeException(nameof(MaxSignedReferences));
        if (MaxCertificates <= 0) throw new ArgumentOutOfRangeException(nameof(MaxCertificates));
        if (MaxCertificateBytes <= 0) throw new ArgumentOutOfRangeException(nameof(MaxCertificateBytes));
        if (MaxTotalCertificateBytes <= 0) throw new ArgumentOutOfRangeException(nameof(MaxTotalCertificateBytes));
    }
}

/// <summary>One signed package-part reference declared by an OPC XML signature.</summary>
public sealed class OfficePackageSignatureReferenceInfo {
    internal OfficePackageSignatureReferenceInfo(string? uri, string? digestMethod, string? digestValue,
        string? targetPartUri, bool? targetPartExists, IReadOnlyList<string> transforms,
        OfficePackageSignatureValidationState digestStatus, string? detail) {
        Uri = uri;
        DigestMethodAlgorithm = digestMethod;
        DigestValue = digestValue;
        TargetPartUri = targetPartUri;
        TargetPartExists = targetPartExists;
        TransformAlgorithms = transforms;
        DigestVerificationStatus = digestStatus;
        DigestVerificationDetail = detail;
    }

    /// <summary>Raw XML DSig Reference URI.</summary>
    public string? Uri { get; }
    /// <summary>Declared digest algorithm.</summary>
    public string? DigestMethodAlgorithm { get; }
    /// <summary>Declared base64 digest value.</summary>
    public string? DigestValue { get; }
    /// <summary>Normalized OPC target part URI, when applicable.</summary>
    public string? TargetPartUri { get; }
    /// <summary>Whether the referenced package part exists.</summary>
    public bool? TargetPartExists { get; }
    /// <summary>Declared transform algorithms in order.</summary>
    public IReadOnlyList<string> TransformAlgorithms { get; }
    /// <summary>Transform-aware package-part digest result.</summary>
    public OfficePackageSignatureValidationState DigestVerificationStatus { get; }
    /// <summary>Deterministic digest detail.</summary>
    public string? DigestVerificationDetail { get; }
    /// <summary>Whether the reference resolves to an OPC package part.</summary>
    public bool IsPackagePartReference => TargetPartUri != null;
}

/// <summary>Timestamp declaration carried by an OPC XML signature.</summary>
public sealed class OfficePackageSignatureTimestampInfo {
    internal OfficePackageSignatureTimestampInfo(string kind, string? value, string? format) {
        Kind = kind;
        Value = value;
        Format = format;
    }

    /// <summary>Timestamp element kind, such as SignatureTime, SigningTime, or EncapsulatedTimeStamp.</summary>
    public string Kind { get; }
    /// <summary>Declared timestamp value when it is a bounded textual time value.</summary>
    public string? Value { get; }
    /// <summary>Declared timestamp format when supplied by the package.</summary>
    public string? Format { get; }
}

/// <summary>One XML signature part discovered in an OPC package.</summary>
public sealed class OfficePackageSignaturePartInfo {
    private readonly IReadOnlyList<byte[]> _certificateBytes;

    internal OfficePackageSignaturePartInfo(string uri, long length, bool isReachableFromOrigin, string? signatureMethod,
        IReadOnlyList<OfficePackageSignatureReferenceInfo> references,
        IReadOnlyList<OfficePackageSignatureTimestampInfo> timestamps, IReadOnlyList<string> subjects,
        IReadOnlyList<byte[]> certificateBytes, string? parseError) {
        Uri = uri;
        Length = length;
        IsReachableFromOrigin = isReachableFromOrigin;
        SignatureMethodAlgorithm = signatureMethod;
        SignedReferences = references;
        Timestamps = timestamps;
        X509SubjectNames = subjects;
        _certificateBytes = certificateBytes;
        ParseError = parseError;
    }

    /// <summary>Signature part URI.</summary>
    public string Uri { get; }
    /// <summary>Encoded signature part length.</summary>
    public long Length { get; }
    /// <summary>Whether the part is reached through the unique internal root-origin and origin-signature relationship chain.</summary>
    public bool IsReachableFromOrigin { get; }
    /// <summary>Declared XML DSig signature method.</summary>
    public string? SignatureMethodAlgorithm { get; }
    /// <summary>Signed package-part references from the OPC Manifest.</summary>
    public IReadOnlyList<OfficePackageSignatureReferenceInfo> SignedReferences { get; }
    /// <summary>Recognized timestamp declarations.</summary>
    public IReadOnlyList<OfficePackageSignatureTimestampInfo> Timestamps { get; }
    /// <summary>Recognized timestamp declaration kinds.</summary>
    public IReadOnlyList<string> TimestampKinds => Timestamps.Select(timestamp => timestamp.Kind).ToArray();
    /// <summary>Declared X.509 subject names.</summary>
    public IReadOnlyList<string> X509SubjectNames { get; }
    /// <summary>XML parse or resource-policy error.</summary>
    public string? ParseError { get; }
    /// <summary>Whether the signature part failed structural parsing.</summary>
    public bool HasParseError => !string.IsNullOrWhiteSpace(ParseError);
    internal IReadOnlyList<byte[]> CertificateBytes => _certificateBytes;
}

/// <summary>Dependency-light structural inspection of one OPC package signature carrier.</summary>
public sealed class OfficePackageSignatureInfo {
    internal OfficePackageSignatureInfo(int originRelationshipCount, int originPartCount, string? originUri, bool hasApplicationMetadata,
        bool signatureDiscoveryComplete, IReadOnlyList<OfficePackageSignaturePartInfo> parts, IReadOnlyList<string> findings) {
        OriginRelationshipCount = originRelationshipCount;
        OriginPartCount = originPartCount;
        HasDigitalSignatureOriginPart = originPartCount > 0;
        OriginPartUri = originUri;
        HasApplicationSignatureMetadata = hasApplicationMetadata;
        SignatureDiscoveryComplete = signatureDiscoveryComplete;
        SignatureParts = parts;
        Findings = findings;
    }

    /// <summary>Number of root digital-signature-origin relationships.</summary>
    public int OriginRelationshipCount { get; }
    /// <summary>Number of distinct existing parts targeted by internal origin relationships.</summary>
    public int OriginPartCount { get; }
    /// <summary>Whether the root relationship and origin part are present.</summary>
    public bool HasDigitalSignatureOriginPart { get; }
    /// <summary>Resolved signature-origin part URI.</summary>
    public string? OriginPartUri { get; }
    /// <summary>Whether extended application properties advertise signatures.</summary>
    public bool HasApplicationSignatureMetadata { get; }
    /// <summary>Whether all signature parts were discovered within the configured limit.</summary>
    public bool SignatureDiscoveryComplete { get; }
    /// <summary>XML signature parts discovered by content type or signature relationships.</summary>
    public IReadOnlyList<OfficePackageSignaturePartInfo> SignatureParts { get; }
    /// <summary>Stable structural findings.</summary>
    public IReadOnlyList<string> Findings { get; }
    /// <summary>Whether any package signature carrier evidence exists.</summary>
    public bool HasSignatures => HasDigitalSignatureOriginPart || HasApplicationSignatureMetadata || SignatureParts.Count > 0;
}

/// <summary>Caller policy for cryptographic OPC signature validation.</summary>
public sealed class OfficePackageSignatureValidationOptions {
    /// <summary>Structural and resource policy.</summary>
    public OfficePackageSignatureInspectionOptions Inspection { get; } = new OfficePackageSignatureInspectionOptions();
    /// <summary>Signer certificate trust and revocation policy.</summary>
    public CertificateValidationOptions CertificateValidation { get; } = new CertificateValidationOptions();
    /// <summary>Whether signer certificate trust is required. Defaults to true.</summary>
    public bool ValidateCertificateTrust { get; set; } = true;
}

/// <summary>Provider-backed cryptographic result for one OPC XML signature part.</summary>
public sealed class OfficePackageSignaturePartValidationResult {
    internal OfficePackageSignaturePartValidationResult(OfficePackageSignaturePartInfo part,
        OfficePackageSignatureValidationState cryptographicStatus,
        OfficePackageSignatureValidationState certificateChainStatus,
        OfficePackageSignatureValidationState revocationStatus,
        bool certificateTrustRequired,
        bool revocationRequired,
        IReadOnlyList<SecurityFinding> findings) {
        SignaturePart = part;
        CryptographicStatus = cryptographicStatus;
        CertificateChainStatus = certificateChainStatus;
        RevocationStatus = revocationStatus;
        CertificateTrustRequired = certificateTrustRequired;
        RevocationRequired = revocationRequired;
        Findings = findings;
    }

    /// <summary>Structurally parsed signature part.</summary>
    public OfficePackageSignaturePartInfo SignaturePart { get; }
    /// <summary>XML signature math and signed-object result.</summary>
    public OfficePackageSignatureValidationState CryptographicStatus { get; }
    /// <summary>Signer certificate-chain result.</summary>
    public OfficePackageSignatureValidationState CertificateChainStatus { get; }
    /// <summary>Signer revocation result.</summary>
    public OfficePackageSignatureValidationState RevocationStatus { get; }
    /// <summary>Whether caller policy required certificate-chain validation.</summary>
    public bool CertificateTrustRequired { get; }
    /// <summary>Whether caller policy required a conclusive revocation result.</summary>
    public bool RevocationRequired { get; }
    /// <summary>Provider and policy findings.</summary>
    public IReadOnlyList<SecurityFinding> Findings { get; }
    /// <summary>Whether this signature satisfies package digests, signature math, and trust policy.</summary>
    public bool IsValidUnderPolicy =>
        SignaturePart.IsReachableFromOrigin && !SignaturePart.HasParseError && SignaturePart.SignedReferences.Count > 0 &&
        SignaturePart.SignedReferences.All(reference => reference.IsPackagePartReference &&
            reference.TargetPartExists == true &&
            reference.DigestVerificationStatus == OfficePackageSignatureValidationState.Passed) &&
        CryptographicStatus == OfficePackageSignatureValidationState.Passed &&
        (CertificateChainStatus == OfficePackageSignatureValidationState.Passed ||
            (!CertificateTrustRequired && CertificateChainStatus == OfficePackageSignatureValidationState.NotChecked)) &&
        (RevocationStatus == OfficePackageSignatureValidationState.Passed ||
            (!RevocationRequired && RevocationStatus != OfficePackageSignatureValidationState.Failed));
}

/// <summary>Combined structural and provider-backed OPC signature validation report.</summary>
public sealed class OfficePackageSignatureValidationReport {
    internal OfficePackageSignatureValidationReport(OfficePackageSignatureInfo signatureInfo,
        IReadOnlyList<OfficePackageSignaturePartValidationResult> signatures,
        IReadOnlyList<string> findings) {
        SignatureInfo = signatureInfo;
        Signatures = signatures;
        Findings = findings;
    }

    /// <summary>Structural package evidence.</summary>
    public OfficePackageSignatureInfo SignatureInfo { get; }
    /// <summary>Per-signature cryptographic and trust evidence.</summary>
    public IReadOnlyList<OfficePackageSignaturePartValidationResult> Signatures { get; }
    /// <summary>Combined deterministic findings.</summary>
    public IReadOnlyList<string> Findings { get; }
    /// <summary>Whether signatures exist and every discovered signature satisfies policy.</summary>
    public bool IsValidUnderPolicy => SignatureInfo.SignatureDiscoveryComplete && SignatureInfo.HasSignatures && Signatures.Count > 0 &&
        Signatures.All(signature => signature.IsValidUnderPolicy);

    /// <summary>Whether package digests and XML signature math pass, independently of certificate trust.</summary>
    public bool IsCryptographicallyValid => SignatureInfo.SignatureDiscoveryComplete && SignatureInfo.HasSignatures && Signatures.Count > 0 &&
        Signatures.All(signature => signature.SignaturePart.IsReachableFromOrigin && !signature.SignaturePart.HasParseError &&
            signature.SignaturePart.SignedReferences.Count > 0 &&
            signature.SignaturePart.SignedReferences.All(reference =>
                reference.IsPackagePartReference && reference.TargetPartExists == true &&
                reference.DigestVerificationStatus == OfficePackageSignatureValidationState.Passed) &&
            signature.CryptographicStatus == OfficePackageSignatureValidationState.Passed);
}
