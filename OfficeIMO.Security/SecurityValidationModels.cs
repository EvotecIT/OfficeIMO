using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;

namespace OfficeIMO.Security;

/// <summary>Validation state shared by CMS, timestamp, and certificate operations.</summary>
public enum SecurityValidationStatus {
    /// <summary>The check was intentionally not performed.</summary>
    NotPerformed = 0,
    /// <summary>The check completed and accepted the supplied evidence.</summary>
    Valid = 1,
    /// <summary>The check completed and rejected the supplied evidence.</summary>
    Invalid = 2,
    /// <summary>The check could not produce a definitive answer.</summary>
    Indeterminate = 3
}

/// <summary>Severity of a structured security finding.</summary>
public enum SecurityFindingSeverity {
    /// <summary>Informational evidence.</summary>
    Info = 0,
    /// <summary>A condition that needs caller attention but does not by itself prove invalidity.</summary>
    Warning = 1,
    /// <summary>A validation or processing failure.</summary>
    Error = 2
}

/// <summary>Stable diagnostic emitted by the neutral security engine.</summary>
public sealed class SecurityFinding {
    /// <summary>Creates a finding.</summary>
    public SecurityFinding(SecurityFindingSeverity severity, string code, string message, int? signerIndex = null) {
        if (string.IsNullOrWhiteSpace(code)) throw new ArgumentException("Finding code cannot be empty.", nameof(code));
        Severity = severity;
        Code = code;
        Message = message ?? string.Empty;
        SignerIndex = signerIndex;
    }

    /// <summary>Finding severity.</summary>
    public SecurityFindingSeverity Severity { get; }
    /// <summary>Stable machine-readable code.</summary>
    public string Code { get; }
    /// <summary>Human-readable explanation.</summary>
    public string Message { get; }
    /// <summary>Zero-based CMS signer index when the finding belongs to one signer.</summary>
    public int? SignerIndex { get; }
}

/// <summary>Caller-controlled platform X.509 chain and trust policy.</summary>
public sealed class CertificateValidationOptions {
    /// <summary>Whether to build a platform certificate chain. Defaults to true.</summary>
    public bool ValidateChain { get; set; } = true;
    /// <summary>Revocation mode. Defaults to NoCheck so validation never silently starts network retrieval.</summary>
    public X509RevocationMode RevocationMode { get; set; } = X509RevocationMode.NoCheck;
    /// <summary>Certificate portion covered by revocation checking.</summary>
    public X509RevocationFlag RevocationFlag { get; set; } = X509RevocationFlag.ExcludeRoot;
    /// <summary>Additional platform verification flags.</summary>
    public X509VerificationFlags VerificationFlags { get; set; } = X509VerificationFlags.NoFlag;
    /// <summary>Optional verification time. The current system time is used when omitted.</summary>
    public DateTime? VerificationTime { get; set; }
    /// <summary>Maximum platform URL retrieval duration when revocation policy permits retrieval.</summary>
    public TimeSpan UrlRetrievalTimeout { get; set; } = TimeSpan.FromSeconds(15);
    /// <summary>Additional intermediate, root, signer, TSA, or recipient certificates.</summary>
    public X509Certificate2Collection ExtraCertificates { get; } = new X509Certificate2Collection();
    /// <summary>Optional application trust callback. Return true to accept the built chain under caller policy.</summary>
    public Func<X509Certificate2, X509Chain, bool>? ChainEvaluator { get; set; }
}

/// <summary>Certificate chain and revocation outcome.</summary>
public sealed class CertificateValidationResult {
    internal CertificateValidationResult(
        SecurityValidationStatus chainStatus,
        SecurityValidationStatus revocationStatus,
        IReadOnlyList<string> chainStatuses) {
        ChainStatus = chainStatus;
        RevocationStatus = revocationStatus;
        ChainStatuses = chainStatuses;
    }

    /// <summary>Platform chain and caller trust-policy outcome.</summary>
    public SecurityValidationStatus ChainStatus { get; }
    /// <summary>Revocation outcome, kept separate from general chain trust.</summary>
    public SecurityValidationStatus RevocationStatus { get; }
    /// <summary>Platform chain status names and messages.</summary>
    public IReadOnlyList<string> ChainStatuses { get; }
}

/// <summary>Options for CMS SignedData verification.</summary>
public sealed class CmsVerificationOptions {
    /// <summary>Maximum encoded CMS bytes accepted. Defaults to 128 MiB.</summary>
    public long MaxEncodedBytes { get; set; } = 128L * 1024 * 1024;
    /// <summary>Maximum detached or encapsulated content bytes accepted. Defaults to 512 MiB.</summary>
    public long MaxContentBytes { get; set; } = 512L * 1024 * 1024;
    /// <summary>Maximum signer count. Defaults to 32.</summary>
    public int MaxSigners { get; set; } = 32;
    /// <summary>Maximum embedded certificate count. Defaults to 256.</summary>
    public int MaxCertificates { get; set; } = 256;
    /// <summary>Whether signature timestamp tokens should be verified. Defaults to true.</summary>
    public bool ValidateTimestamps { get; set; } = true;
    /// <summary>Certificate-chain policy shared by signers and timestamp authorities.</summary>
    public CertificateValidationOptions CertificateValidation { get; } = new CertificateValidationOptions();
}

/// <summary>Options for CMS signing.</summary>
public sealed class CmsSigningOptions {
    /// <summary>Digest algorithm. SHA-256 is the default.</summary>
    public HashAlgorithmName DigestAlgorithm { get; set; } = HashAlgorithmName.SHA256;
    /// <summary>Whether to include the CMS signing-time attribute. Defaults to true.</summary>
    public bool IncludeSigningTime { get; set; } = true;
    /// <summary>Optional signing time. Current UTC time is used when omitted.</summary>
    public DateTimeOffset? SigningTime { get; set; }
    /// <summary>Whether supplied chain certificates are embedded. Defaults to true.</summary>
    public bool IncludeCertificateChain { get; set; } = true;
    /// <summary>Maximum content bytes accepted. Defaults to 512 MiB.</summary>
    public long MaxContentBytes { get; set; } = 512L * 1024 * 1024;
}

/// <summary>One signer from a CMS SignedData object.</summary>
public sealed class CmsSignerVerificationResult {
    internal CmsSignerVerificationResult(
        int signerIndex,
        SecurityValidationStatus signatureStatus,
        SecurityValidationStatus digestStatus,
        CertificateValidationResult certificateValidation,
        SecurityValidationStatus timestampStatus,
        byte[]? signerCertificate,
        string? subject,
        string? issuer,
        string? serialNumber,
        string? thumbprint,
        string digestAlgorithmOid,
        string signatureAlgorithmOid,
        DateTimeOffset? signingTime,
        DateTimeOffset? timestampTime,
        IReadOnlyList<Rfc3161TimestampVerificationResult> timestampTokens,
        IReadOnlyList<SecurityFinding> findings) {
        SignerIndex = signerIndex;
        SignatureStatus = signatureStatus;
        DigestStatus = digestStatus;
        CertificateValidation = certificateValidation;
        TimestampStatus = timestampStatus;
        SignerCertificate = signerCertificate;
        Subject = subject;
        Issuer = issuer;
        SerialNumber = serialNumber;
        Thumbprint = thumbprint;
        DigestAlgorithmOid = digestAlgorithmOid;
        SignatureAlgorithmOid = signatureAlgorithmOid;
        SigningTime = signingTime;
        TimestampTime = timestampTime;
        TimestampTokens = timestampTokens;
        Findings = findings;
    }

    /// <summary>Zero-based signer index.</summary>
    public int SignerIndex { get; }
    /// <summary>Mathematical signature and signed-attribute outcome.</summary>
    public SecurityValidationStatus SignatureStatus { get; }
    /// <summary>Content digest binding outcome.</summary>
    public SecurityValidationStatus DigestStatus { get; }
    /// <summary>Signer certificate trust and revocation outcome.</summary>
    public CertificateValidationResult CertificateValidation { get; }
    /// <summary>Signature timestamp-token outcome.</summary>
    public SecurityValidationStatus TimestampStatus { get; }
    /// <summary>DER signer certificate bytes, when available.</summary>
    public byte[]? SignerCertificate { get; }
    /// <summary>Signer certificate subject.</summary>
    public string? Subject { get; }
    /// <summary>Signer certificate issuer.</summary>
    public string? Issuer { get; }
    /// <summary>Signer certificate serial number.</summary>
    public string? SerialNumber { get; }
    /// <summary>Signer certificate thumbprint.</summary>
    public string? Thumbprint { get; }
    /// <summary>CMS digest algorithm object identifier.</summary>
    public string DigestAlgorithmOid { get; }
    /// <summary>CMS signature algorithm object identifier.</summary>
    public string SignatureAlgorithmOid { get; }
    /// <summary>CMS signing time, when present and validly encoded.</summary>
    public DateTimeOffset? SigningTime { get; }
    /// <summary>Validated signature timestamp time, when present.</summary>
    public DateTimeOffset? TimestampTime { get; }
    /// <summary>Every signature timestamp token found on this signer.</summary>
    public IReadOnlyList<Rfc3161TimestampVerificationResult> TimestampTokens { get; }
    /// <summary>Structured findings scoped to this signer.</summary>
    public IReadOnlyList<SecurityFinding> Findings { get; }
}

/// <summary>Neutral CMS SignedData verification result.</summary>
public sealed class CmsVerificationResult {
    internal CmsVerificationResult(
        bool parsed,
        bool isDetached,
        string? contentTypeOid,
        byte[]? encapsulatedContent,
        IReadOnlyList<CmsSignerVerificationResult> signers,
        IReadOnlyList<SecurityFinding> findings) {
        Parsed = parsed;
        IsDetached = isDetached;
        ContentTypeOid = contentTypeOid;
        EncapsulatedContent = encapsulatedContent;
        Signers = signers;
        Findings = findings;
    }

    /// <summary>Whether the CMS container was decoded.</summary>
    public bool Parsed { get; }
    /// <summary>Whether SignedData omits encapsulated content.</summary>
    public bool IsDetached { get; }
    /// <summary>Encapsulated content type object identifier.</summary>
    public string? ContentTypeOid { get; }
    /// <summary>Cloned encapsulated content, when present.</summary>
    public byte[]? EncapsulatedContent { get; }
    /// <summary>Signer results in encoded order.</summary>
    public IReadOnlyList<CmsSignerVerificationResult> Signers { get; }
    /// <summary>Container-level findings.</summary>
    public IReadOnlyList<SecurityFinding> Findings { get; }
    /// <summary>True only when at least one signer exists and every signer signature and digest is valid.</summary>
    public bool IsCryptographicallyValid => Signers.Count > 0 &&
        Signers.All(static signer => signer.SignatureStatus == SecurityValidationStatus.Valid &&
            signer.DigestStatus == SecurityValidationStatus.Valid);
}

/// <summary>RFC 3161 timestamp-token verification result.</summary>
public sealed class Rfc3161TimestampVerificationResult {
    internal Rfc3161TimestampVerificationResult(
        SecurityValidationStatus status,
        DateTimeOffset? timestamp,
        string? policyOid,
        string? messageImprintAlgorithmOid,
        byte[]? tsaCertificate,
        CertificateValidationResult certificateValidation,
        IReadOnlyList<SecurityFinding> findings) {
        Status = status;
        Timestamp = timestamp;
        PolicyOid = policyOid;
        MessageImprintAlgorithmOid = messageImprintAlgorithmOid;
        TsaCertificate = tsaCertificate;
        CertificateValidation = certificateValidation;
        Findings = findings;
    }

    /// <summary>Combined token signature, message-imprint, and TSA certificate-validation outcome.</summary>
    public SecurityValidationStatus Status { get; }
    /// <summary>Timestamp generation time.</summary>
    public DateTimeOffset? Timestamp { get; }
    /// <summary>Timestamp policy object identifier.</summary>
    public string? PolicyOid { get; }
    /// <summary>Message-imprint digest algorithm object identifier.</summary>
    public string? MessageImprintAlgorithmOid { get; }
    /// <summary>DER TSA signer certificate bytes.</summary>
    public byte[]? TsaCertificate { get; }
    /// <summary>TSA certificate trust and revocation outcome.</summary>
    public CertificateValidationResult CertificateValidation { get; }
    /// <summary>Structured findings.</summary>
    public IReadOnlyList<SecurityFinding> Findings { get; }
}
