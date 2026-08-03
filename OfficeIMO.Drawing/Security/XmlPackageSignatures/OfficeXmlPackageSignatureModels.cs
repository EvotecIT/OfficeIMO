using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;

namespace OfficeIMO.Security;

/// <summary>Supported ZIP-container XML signature profiles.</summary>
public enum OfficeXmlPackageSignatureFormat {
    /// <summary>ODF META-INF/documentsignatures.xml.</summary>
    OpenDocument,
    /// <summary>EPUB OCF META-INF/signatures.xml.</summary>
    Epub
}

/// <summary>Bounded creation and validation policy for ODF/EPUB XML signatures.</summary>
public sealed class OfficeXmlPackageSignatureOptions {
    /// <summary>Maximum package bytes. Defaults to 512 MiB.</summary>
    public long MaxPackageBytes { get; set; } = 512L * 1024L * 1024L;
    /// <summary>Maximum ZIP entries. Defaults to 10,000.</summary>
    public int MaxEntries { get; set; } = 10000;
    /// <summary>Maximum bytes read from one entry. Defaults to 256 MiB.</summary>
    public long MaxEntryBytes { get; set; } = 256L * 1024L * 1024L;
    /// <summary>Maximum aggregate bytes hashed. Defaults to 512 MiB.</summary>
    public long MaxTotalDigestBytes { get; set; } = 512L * 1024L * 1024L;
    /// <summary>Maximum signature carrier bytes. Defaults to 16 MiB.</summary>
    public long MaxSignatureBytes { get; set; } = 16L * 1024L * 1024L;
    /// <summary>Maximum signature count. Defaults to 32.</summary>
    public int MaxSignatures { get; set; } = 32;
    /// <summary>Maximum embedded certificates per signature. Defaults to 64.</summary>
    public int MaxCertificates { get; set; } = 64;
    /// <summary>Maximum decoded bytes per embedded certificate. Defaults to 4 MiB.</summary>
    public long MaxCertificateBytes { get; set; } = 4L * 1024L * 1024L;
    /// <summary>Maximum aggregate decoded certificate bytes per signature. Defaults to 64 MiB.</summary>
    public long MaxTotalCertificateBytes { get; set; } = 64L * 1024L * 1024L;
    /// <summary>Signer certificate trust and revocation policy.</summary>
    public CertificateValidationOptions CertificateValidation { get; } = new CertificateValidationOptions();
    /// <summary>Whether signer trust must be validated. Defaults to true.</summary>
    public bool ValidateCertificateTrust { get; set; } = true;
    /// <summary>Additional certificates embedded after the signer certificate.</summary>
    public IReadOnlyCollection<X509Certificate2>? AdditionalCertificates { get; set; }
}

/// <summary>One package entry digest declared by a signed XML manifest.</summary>
public sealed class OfficeXmlPackageEntryDigestResult {
    internal OfficeXmlPackageEntryDigestResult(string path, bool exists,
        OfficePackageSignatureValidationState status, string detail) {
        Path = path; Exists = exists; Status = status; Detail = detail;
    }
    /// <summary>Normalized ZIP entry path.</summary>
    public string Path { get; }
    /// <summary>Whether the entry exists.</summary>
    public bool Exists { get; }
    /// <summary>Digest validation result.</summary>
    public OfficePackageSignatureValidationState Status { get; }
    /// <summary>Deterministic detail.</summary>
    public string Detail { get; }
}

/// <summary>Provider and package-digest validation for one ODF/EPUB XML signature.</summary>
public sealed class OfficeXmlPackageSignatureResult {
    internal OfficeXmlPackageSignatureResult(string? signatureId,
        OfficePackageSignatureValidationState cryptographicStatus,
        OfficePackageSignatureValidationState certificateChainStatus,
        OfficePackageSignatureValidationState revocationStatus,
        bool certificateTrustRequired,
        IReadOnlyList<OfficeXmlPackageEntryDigestResult> entries,
        IReadOnlyList<SecurityFinding> findings) {
        SignatureId = signatureId;
        CryptographicStatus = cryptographicStatus;
        CertificateChainStatus = certificateChainStatus;
        RevocationStatus = revocationStatus;
        CertificateTrustRequired = certificateTrustRequired;
        Entries = entries;
        Findings = findings;
    }
    /// <summary>XML Signature Id.</summary>
    public string? SignatureId { get; }
    /// <summary>XML signature math result.</summary>
    public OfficePackageSignatureValidationState CryptographicStatus { get; }
    /// <summary>Signer certificate-chain result.</summary>
    public OfficePackageSignatureValidationState CertificateChainStatus { get; }
    /// <summary>Signer revocation result.</summary>
    public OfficePackageSignatureValidationState RevocationStatus { get; }
    /// <summary>Whether caller policy required certificate-chain validation.</summary>
    public bool CertificateTrustRequired { get; }
    /// <summary>Signed package-entry digest results.</summary>
    public IReadOnlyList<OfficeXmlPackageEntryDigestResult> Entries { get; }
    /// <summary>Provider findings.</summary>
    public IReadOnlyList<SecurityFinding> Findings { get; }
    /// <summary>Whether signature math, all package digests, and requested trust policy pass.</summary>
    public bool IsValidUnderPolicy => CryptographicStatus == OfficePackageSignatureValidationState.Passed &&
        Entries.Count > 0 && Entries.All(entry => entry.Status == OfficePackageSignatureValidationState.Passed) &&
        (CertificateChainStatus == OfficePackageSignatureValidationState.Passed ||
            (!CertificateTrustRequired && CertificateChainStatus == OfficePackageSignatureValidationState.NotChecked)) &&
        RevocationStatus != OfficePackageSignatureValidationState.Failed;
}

/// <summary>Combined XML signature report for one ODF or EPUB package.</summary>
public sealed class OfficeXmlPackageSignatureValidationReport {
    internal OfficeXmlPackageSignatureValidationReport(string filePath, string carrierPath,
        bool carrierPresent, bool carrierWellFormed,
        IReadOnlyList<OfficeXmlPackageSignatureResult> signatures,
        IReadOnlyList<string> findings) {
        FilePath = filePath; CarrierPath = carrierPath; CarrierPresent = carrierPresent;
        CarrierWellFormed = carrierWellFormed; Signatures = signatures; Findings = findings;
    }
    /// <summary>Normalized package path.</summary>
    public string FilePath { get; }
    /// <summary>Format-defined signature carrier path.</summary>
    public string CarrierPath { get; }
    /// <summary>Whether the carrier exists.</summary>
    public bool CarrierPresent { get; }
    /// <summary>Whether the carrier parsed within limits.</summary>
    public bool CarrierWellFormed { get; }
    /// <summary>Individual signature validation results.</summary>
    public IReadOnlyList<OfficeXmlPackageSignatureResult> Signatures { get; }
    /// <summary>Package-level findings.</summary>
    public IReadOnlyList<string> Findings { get; }
    /// <summary>Whether at least one signature exists and every signature satisfies policy.</summary>
    public bool IsValidUnderPolicy => CarrierPresent && CarrierWellFormed && Signatures.Count > 0 &&
        Signatures.All(signature => signature.IsValidUnderPolicy);
}

/// <summary>Atomic XML package signature creation result.</summary>
public sealed class OfficeXmlPackageSigningResult {
    internal OfficeXmlPackageSigningResult(string filePath, bool succeeded, int signatureCount,
        OfficeXmlPackageSignatureValidationReport? validation, IReadOnlyList<string> findings) {
        FilePath = filePath; Succeeded = succeeded; SignatureCount = signatureCount;
        Validation = validation; Findings = findings;
    }
    /// <summary>Package path.</summary>
    public string FilePath { get; }
    /// <summary>Whether creation, validation readback, and atomic commit succeeded.</summary>
    public bool Succeeded { get; }
    /// <summary>Signature count after creation.</summary>
    public int SignatureCount { get; }
    /// <summary>Validation readback.</summary>
    public OfficeXmlPackageSignatureValidationReport? Validation { get; }
    /// <summary>Creation findings.</summary>
    public IReadOnlyList<string> Findings { get; }
}
