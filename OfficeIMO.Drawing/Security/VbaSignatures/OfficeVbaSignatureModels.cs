using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;

namespace OfficeIMO.Security;

/// <summary>Microsoft Office VBA signature profile carried next to vbaProject.bin.</summary>
public enum OfficeVbaSignatureProfile {
    /// <summary>Office 2007-compatible signature carrier.</summary>
    Legacy = 1,
    /// <summary>Agile Office 2013+ signature carrier.</summary>
    Agile = 2,
    /// <summary>V3 Office signature carrier.</summary>
    V3 = 3
}

/// <summary>Stable VBA signature finding.</summary>
public sealed class OfficeVbaSignatureFinding {
    internal OfficeVbaSignatureFinding(string code, OfficePackageSignatureValidationState state,
        string message, OfficeVbaSignatureProfile? profile = null) {
        Code = code;
        State = state;
        Message = message;
        Profile = profile;
    }
    /// <summary>Machine-readable code.</summary>
    public string Code { get; }
    /// <summary>Finding state.</summary>
    public OfficePackageSignatureValidationState State { get; }
    /// <summary>Human-readable detail.</summary>
    public string Message { get; }
    /// <summary>Related profile, when applicable.</summary>
    public OfficeVbaSignatureProfile? Profile { get; }
}

/// <summary>One VBA signature profile part.</summary>
public sealed class OfficeVbaSignaturePartInfo {
    internal OfficeVbaSignaturePartInfo(OfficeVbaSignatureProfile profile, string uri,
        string relationshipType, string contentType, long length, bool cmsParsed,
        OfficePackageSignatureValidationState cryptographicStatus,
        OfficePackageSignatureValidationState certificateChainStatus,
        OfficePackageSignatureValidationState revocationStatus,
        OfficePackageSignatureValidationState timestampStatus,
        string? signerSubject, string? signerThumbprint,
        string? subjectDigestAlgorithmOid, byte[]? subjectDigest,
        IReadOnlyList<OfficeVbaSignatureFinding> findings) {
        Profile = profile;
        Uri = uri;
        RelationshipType = relationshipType;
        ContentType = contentType;
        Length = length;
        CmsParsed = cmsParsed;
        CryptographicStatus = cryptographicStatus;
        CertificateChainStatus = certificateChainStatus;
        RevocationStatus = revocationStatus;
        TimestampStatus = timestampStatus;
        SignerSubject = signerSubject;
        SignerThumbprint = signerThumbprint;
        SubjectDigestAlgorithmOid = subjectDigestAlgorithmOid;
        SubjectDigest = subjectDigest == null ? null : (byte[])subjectDigest.Clone();
        Findings = findings;
    }
    /// <summary>Signature profile.</summary>
    public OfficeVbaSignatureProfile Profile { get; }
    /// <summary>Package part URI.</summary>
    public string Uri { get; }
    /// <summary>Relationship type that selected the profile.</summary>
    public string RelationshipType { get; }
    /// <summary>Declared content type.</summary>
    public string ContentType { get; }
    /// <summary>Encoded part length.</summary>
    public long Length { get; }
    /// <summary>Whether the bounded DigSigInfoSerialized wrapper exposed CMS bytes.</summary>
    public bool CmsParsed { get; }
    /// <summary>CMS signature and digest result.</summary>
    public OfficePackageSignatureValidationState CryptographicStatus { get; }
    /// <summary>Signer certificate-chain result.</summary>
    public OfficePackageSignatureValidationState CertificateChainStatus { get; }
    /// <summary>Signer revocation result.</summary>
    public OfficePackageSignatureValidationState RevocationStatus { get; }
    /// <summary>RFC 3161 timestamp result.</summary>
    public OfficePackageSignatureValidationState TimestampStatus { get; }
    /// <summary>Signer subject.</summary>
    public string? SignerSubject { get; }
    /// <summary>Signer thumbprint.</summary>
    public string? SignerThumbprint { get; }
    /// <summary>Authenticode Office SIP digest algorithm.</summary>
    public string? SubjectDigestAlgorithmOid { get; }
    /// <summary>Cloned Authenticode Office SIP digest.</summary>
    public byte[]? SubjectDigest { get; }
    /// <summary>Profile findings.</summary>
    public IReadOnlyList<OfficeVbaSignatureFinding> Findings { get; }
}

/// <summary>Bounded VBA project and signature profile evidence for a saved Office package.</summary>
public sealed class OfficeVbaSignatureInfo {
    internal OfficeVbaSignatureInfo(string filePath, bool isMacroEnabledFormat, bool hasMacroProject,
        string? macroProjectUri, long? macroProjectLength, string? macroProjectSha256,
        IReadOnlyList<OfficeVbaSignaturePartInfo> signatures,
        IReadOnlyList<OfficeVbaSignatureFinding> findings) {
        FilePath = filePath;
        IsMacroEnabledFormat = isMacroEnabledFormat;
        HasMacroProject = hasMacroProject;
        MacroProjectUri = macroProjectUri;
        MacroProjectLength = macroProjectLength;
        MacroProjectSha256 = macroProjectSha256;
        Signatures = signatures;
        Findings = findings;
    }
    /// <summary>Normalized package path.</summary>
    public string FilePath { get; }
    /// <summary>Whether the extension is a supported macro-capable Word, Excel, or PowerPoint format.</summary>
    public bool IsMacroEnabledFormat { get; }
    /// <summary>Whether vbaProject.bin exists.</summary>
    public bool HasMacroProject { get; }
    /// <summary>VBA project part URI.</summary>
    public string? MacroProjectUri { get; }
    /// <summary>VBA project byte length.</summary>
    public long? MacroProjectLength { get; }
    /// <summary>SHA-256 of the exact VBA compound storage.</summary>
    public string? MacroProjectSha256 { get; }
    /// <summary>Legacy, agile, and V3 profile evidence.</summary>
    public IReadOnlyList<OfficeVbaSignaturePartInfo> Signatures { get; }
    /// <summary>Package-level findings.</summary>
    public IReadOnlyList<OfficeVbaSignatureFinding> Findings { get; }
    /// <summary>Whether at least one profile is present.</summary>
    public bool HasSignatures => Signatures.Count > 0;
    /// <summary>Whether the V3 profile is present.</summary>
    public bool HasV3Signature => Signatures.Any(signature => signature.Profile == OfficeVbaSignatureProfile.V3);
}

/// <summary>Bounded VBA signature inspection and CMS policy.</summary>
public class OfficeVbaSignatureInspectionOptions {
    /// <summary>Package resource policy.</summary>
    public OfficePackageSignatureInspectionOptions Package { get; } = new OfficePackageSignatureInspectionOptions();
    /// <summary>CMS trust, revocation, and timestamp policy.</summary>
    public CmsVerificationOptions CmsVerification { get; } = new CmsVerificationOptions();
    /// <summary>Maximum VBA project bytes. Defaults to 64 MiB.</summary>
    public long MaxMacroProjectBytes { get; set; } = 64L * 1024L * 1024L;
    /// <summary>Maximum one profile part bytes. Defaults to 32 MiB.</summary>
    public long MaxSignatureBytes { get; set; } = 32L * 1024L * 1024L;
    /// <summary>Maximum aggregate profile bytes. Defaults to 64 MiB.</summary>
    public long MaxTotalSignatureBytes { get; set; } = 64L * 1024L * 1024L;
    /// <summary>Maximum vbaProject relationships. Defaults to 128.</summary>
    public int MaxRelationships { get; set; } = 128;
    /// <summary>Whether CMS, trust, revocation, and timestamps are validated when a provider is supplied.</summary>
    public bool ValidateCms { get; set; } = true;
}

/// <summary>VBA signature validation result including Office SIP content binding.</summary>
public sealed class OfficeVbaSignatureValidationResult {
    internal OfficeVbaSignatureValidationResult(OfficeVbaSignatureInfo info, bool contentBindingSupported,
        OfficePackageSignatureValidationState contentBindingStatus,
        bool revocationRequired,
        IReadOnlyList<OfficeVbaSignatureFinding> findings) {
        SignatureInfo = info;
        ContentBindingSupported = contentBindingSupported;
        ContentBindingStatus = contentBindingStatus;
        RevocationRequired = revocationRequired;
        Findings = findings;
    }
    /// <summary>Structural and CMS profile evidence.</summary>
    public OfficeVbaSignatureInfo SignatureInfo { get; }
    /// <summary>Whether managed MS-OVBA content binding was available for the selected project.</summary>
    public bool ContentBindingSupported { get; }
    /// <summary>Office SIP content-binding result.</summary>
    public OfficePackageSignatureValidationState ContentBindingStatus { get; }
    /// <summary>Whether the supplied CMS policy requires a conclusive revocation result.</summary>
    public bool RevocationRequired { get; }
    /// <summary>Combined findings.</summary>
    public IReadOnlyList<OfficeVbaSignatureFinding> Findings { get; }
    /// <summary>Whether the highest profile binds to the package and satisfies CMS/trust policy.</summary>
    public bool IsValidUnderPolicy {
        get {
            OfficeVbaSignaturePartInfo? selected = SignatureInfo.Signatures.OrderByDescending(signature => signature.Profile).FirstOrDefault();
            return selected != null && !SignatureInfo.Findings.Any(finding => finding.State == OfficePackageSignatureValidationState.Failed) &&
                !selected.Findings.Any(finding => finding.State == OfficePackageSignatureValidationState.Failed) &&
                !Findings.Any(finding => finding.State == OfficePackageSignatureValidationState.Failed) &&
                ContentBindingStatus == OfficePackageSignatureValidationState.Passed &&
                selected.CryptographicStatus == OfficePackageSignatureValidationState.Passed &&
                selected.CertificateChainStatus == OfficePackageSignatureValidationState.Passed &&
                (selected.RevocationStatus == OfficePackageSignatureValidationState.Passed ||
                    (!RevocationRequired && selected.RevocationStatus != OfficePackageSignatureValidationState.Failed)) &&
                selected.TimestampStatus != OfficePackageSignatureValidationState.Failed;
        }
    }
}

/// <summary>Portable managed VBA signing policy.</summary>
public sealed class OfficeVbaSigningOptions : OfficeVbaSignatureInspectionOptions {
    /// <summary>Whether signing may invalidate existing OPC package signatures. Defaults to false.</summary>
    public bool AllowPackageSignatureInvalidation { get; set; }
    /// <summary>CMS signer settings. SHA-256 signs each profile while MS-OVBA defines its profile-specific content transcript.</summary>
    public CmsSigningOptions CmsSigning { get; } = new CmsSigningOptions();
    /// <summary>Whether an available registered Microsoft Office SIP should be used as an additional Windows differential check.</summary>
    public bool ValidateWithWindowsSipWhenAvailable { get; set; }
}

/// <summary>Atomic VBA profile signing result.</summary>
public sealed class OfficeVbaSigningResult {
    internal OfficeVbaSigningResult(string filePath, bool supported, bool succeeded,
        OfficeVbaSignatureValidationResult? validation,
        IReadOnlyList<OfficeVbaSignatureFinding> findings) {
        FilePath = filePath;
        IsSupported = supported;
        Succeeded = succeeded;
        Validation = validation;
        Findings = findings;
    }
    /// <summary>Package path.</summary>
    public string FilePath { get; }
    /// <summary>Whether the managed implementation supports the input format and project profile.</summary>
    public bool IsSupported { get; }
    /// <summary>Whether all profiles were created, validated, and committed atomically.</summary>
    public bool Succeeded { get; }
    /// <summary>Post-signing validation evidence.</summary>
    public OfficeVbaSignatureValidationResult? Validation { get; }
    /// <summary>Signing findings.</summary>
    public IReadOnlyList<OfficeVbaSignatureFinding> Findings { get; }
}
