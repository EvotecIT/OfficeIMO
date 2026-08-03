using System.Security.Cryptography.X509Certificates;
using OfficeIMO.Drawing;
using OfficeIMO.Security;

namespace OfficeIMO.Word {
    /// <summary>Identifies one Microsoft VBA macro-project signature profile.</summary>
    public enum WordMacroProjectSignatureProfile {
        /// <summary>The signature profile could not be identified.</summary>
        Unknown = 0,
        /// <summary>The original VBA signature profile.</summary>
        Legacy = 1,
        /// <summary>The agile VBA signature profile introduced with newer Office versions.</summary>
        Agile = 2,
        /// <summary>The current V3 VBA signature profile with the strongest project binding.</summary>
        V3 = 3
    }

    /// <summary>Stable diagnostic for VBA macro-project signature inspection, validation, or creation.</summary>
    public sealed class WordMacroProjectSignatureFinding {
        internal WordMacroProjectSignatureFinding(
            string code,
            WordSignatureValidationState state,
            string message,
            WordMacroProjectSignatureProfile? profile = null) {
            Code = code;
            State = state;
            Message = message;
            Profile = profile;
        }

        /// <summary>Gets the stable machine-readable diagnostic code.</summary>
        public string Code { get; }

        /// <summary>Gets the validation state represented by the diagnostic.</summary>
        public WordSignatureValidationState State { get; }

        /// <summary>Gets the human-readable diagnostic message.</summary>
        public string Message { get; }

        /// <summary>Gets the signature profile when the diagnostic belongs to one profile.</summary>
        public WordMacroProjectSignatureProfile? Profile { get; }
    }

    /// <summary>Resource policy for cross-platform VBA signature-part inspection.</summary>
    public sealed class WordMacroProjectSignatureInspectionOptions {
        /// <summary>Gets the shared ZIP structure, size, compression, and active-content policy applied before Open XML parsing.</summary>
        public OfficePackageSecurityOptions PackageSecurity { get; } = OfficePackageSecurityOptions.SecureDefaults;

        /// <summary>Gets or sets the maximum encoded VBA project size. Defaults to 256 MiB.</summary>
        public long MaxMacroProjectBytes { get; set; } = 256L * 1024 * 1024;

        /// <summary>Gets or sets the maximum encoded bytes for one VBA signature part. Defaults to 16 MiB.</summary>
        public long MaxSignatureBytes { get; set; } = 16L * 1024 * 1024;

        /// <summary>Gets or sets the maximum aggregate encoded bytes for all VBA signature parts. Defaults to 48 MiB.</summary>
        public long MaxTotalSignatureBytes { get; set; } = 48L * 1024 * 1024;

        /// <summary>Gets or sets the maximum relationships accepted from the VBA project part. Defaults to 32.</summary>
        public int MaxRelationships { get; set; } = 32;

        /// <summary>Gets or sets whether CMS signer, trust, revocation, and timestamp metadata is evaluated. Defaults to true.</summary>
        public bool ValidateCms { get; set; } = true;

        /// <summary>Gets the certificate-chain, revocation, timestamp, and CMS resource policy.</summary>
        public CmsVerificationOptions CmsVerification { get; } = new CmsVerificationOptions {
            MaxEncodedBytes = 16L * 1024 * 1024,
            MaxContentBytes = 16L * 1024 * 1024,
            MaxSigners = 1,
            MaxCertificates = 64,
            MaxTimestampTokens = 8,
            MaxTimestampTokenBytes = 16L * 1024 * 1024,
            MaxTotalTimestampBytes = 32L * 1024 * 1024
        };
    }

    /// <summary>Parsed metadata and policy evidence for one VBA signature part.</summary>
    public sealed class WordMacroProjectSignaturePartInfo {
        private readonly byte[]? _signedContentDigest;

        internal WordMacroProjectSignaturePartInfo(
            WordMacroProjectSignatureProfile profile,
            string uri,
            string relationshipType,
            string contentType,
            long length,
            bool cmsParsed,
            WordSignatureValidationState cryptographicStatus,
            WordSignatureValidationState certificateChainStatus,
            WordSignatureValidationState revocationStatus,
            WordSignatureValidationState timestampStatus,
            string? signerSubject,
            string? signerThumbprint,
            string? digestAlgorithmOid,
            string? signatureAlgorithmOid,
            string? signedContentDigestAlgorithmOid,
            byte[]? signedContentDigest,
            DateTimeOffset? signingTime,
            DateTimeOffset? timestampTime,
            IReadOnlyList<WordMacroProjectSignatureFinding> findings) {
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
            DigestAlgorithmOid = digestAlgorithmOid;
            SignatureAlgorithmOid = signatureAlgorithmOid;
            SignedContentDigestAlgorithmOid = signedContentDigestAlgorithmOid;
            _signedContentDigest = signedContentDigest == null ? null : (byte[])signedContentDigest.Clone();
            SigningTime = signingTime;
            TimestampTime = timestampTime;
            Findings = findings;
        }

        /// <summary>Gets the VBA signature profile.</summary>
        public WordMacroProjectSignatureProfile Profile { get; }

        /// <summary>Gets the package part URI.</summary>
        public string Uri { get; }

        /// <summary>Gets the relationship type from the VBA project part.</summary>
        public string RelationshipType { get; }

        /// <summary>Gets the package content type.</summary>
        public string ContentType { get; }

        /// <summary>Gets the encoded signature-part length.</summary>
        public long Length { get; }

        /// <summary>Gets whether the embedded CMS SignedData container was decoded.</summary>
        public bool CmsParsed { get; }

        /// <summary>Gets CMS signature and embedded-digest status. This does not prove the VBA project content binding.</summary>
        public WordSignatureValidationState CryptographicStatus { get; }

        /// <summary>Gets signer certificate-chain trust under caller policy.</summary>
        public WordSignatureValidationState CertificateChainStatus { get; }

        /// <summary>Gets signer revocation status under caller policy.</summary>
        public WordSignatureValidationState RevocationStatus { get; }

        /// <summary>Gets signature timestamp-authority status under caller policy.</summary>
        public WordSignatureValidationState TimestampStatus { get; }

        /// <summary>Gets the signer certificate subject when available.</summary>
        public string? SignerSubject { get; }

        /// <summary>Gets the signer certificate thumbprint when available.</summary>
        public string? SignerThumbprint { get; }

        /// <summary>Gets the CMS digest-algorithm object identifier when available.</summary>
        public string? DigestAlgorithmOid { get; }

        /// <summary>Gets the CMS signature-algorithm object identifier when available.</summary>
        public string? SignatureAlgorithmOid { get; }

        /// <summary>Gets the Authenticode subject-digest algorithm supplied to the Office SIP.</summary>
        public string? SignedContentDigestAlgorithmOid { get; }

        /// <summary>Gets a clone of the Authenticode subject digest signed by this profile.</summary>
        public byte[]? SignedContentDigest => _signedContentDigest == null ? null : (byte[])_signedContentDigest.Clone();

        /// <summary>Gets the signed CMS signing time when present.</summary>
        public DateTimeOffset? SigningTime { get; }

        /// <summary>Gets the validated RFC 3161 timestamp time when present.</summary>
        public DateTimeOffset? TimestampTime { get; }

        /// <summary>Gets stable diagnostics for this signature profile.</summary>
        public IReadOnlyList<WordMacroProjectSignatureFinding> Findings { get; }
    }

    /// <summary>Cross-platform inventory of VBA project and signature parts in one macro-enabled Word package.</summary>
    public sealed class WordMacroProjectSignatureInfo {
        internal WordMacroProjectSignatureInfo(
            string filePath,
            bool isMacroEnabledFormat,
            bool hasMacroProject,
            string? macroProjectPartUri,
            long? macroProjectLength,
            string? macroProjectSha256,
            IReadOnlyList<WordMacroProjectSignaturePartInfo> signatures,
            IReadOnlyList<WordMacroProjectSignatureFinding> findings) {
            FilePath = filePath;
            IsMacroEnabledFormat = isMacroEnabledFormat;
            HasMacroProject = hasMacroProject;
            MacroProjectPartUri = macroProjectPartUri;
            MacroProjectLength = macroProjectLength;
            MacroProjectSha256 = macroProjectSha256;
            Signatures = signatures;
            Findings = findings;
        }

        /// <summary>Gets the inspected file path.</summary>
        public string FilePath { get; }

        /// <summary>Gets whether the extension is DOCM or DOTM.</summary>
        public bool IsMacroEnabledFormat { get; }

        /// <summary>Gets whether a VBA project part exists.</summary>
        public bool HasMacroProject { get; }

        /// <summary>Gets the VBA project package-part URI when present.</summary>
        public string? MacroProjectPartUri { get; }

        /// <summary>Gets the encoded VBA project length when present.</summary>
        public long? MacroProjectLength { get; }

        /// <summary>Gets a SHA-256 preservation fingerprint of the encoded VBA project when inspection succeeded.</summary>
        public string? MacroProjectSha256 { get; }

        /// <summary>Gets VBA signature parts in profile order.</summary>
        public IReadOnlyList<WordMacroProjectSignaturePartInfo> Signatures { get; }

        /// <summary>Gets stable aggregate inspection diagnostics.</summary>
        public IReadOnlyList<WordMacroProjectSignatureFinding> Findings { get; }

        /// <summary>Gets whether any VBA signature profile is present.</summary>
        public bool HasSignatures => Signatures.Count > 0;

        /// <summary>Gets whether the current secure V3 profile is present.</summary>
        public bool HasV3Signature => Signatures.Any(signature => signature.Profile == WordMacroProjectSignatureProfile.V3);
    }

    /// <summary>Windows Office SIP and validation policy shared by VBA signing and verification.</summary>
    public class WordMacroProjectSignatureValidationOptions {
        /// <summary>Gets or sets whether a V3 signature part is required for acceptance. Defaults to true.</summary>
        public bool RequireV3Signature { get; set; } = true;

        /// <summary>Gets or sets whether a valid RFC 3161 timestamp is required for acceptance. Defaults to false.</summary>
        public bool RequireTimestamp { get; set; }

        /// <summary>Gets the cross-platform signature and certificate inspection policy.</summary>
        public WordMacroProjectSignatureInspectionOptions Inspection { get; } = new WordMacroProjectSignatureInspectionOptions();
    }

    /// <summary>Options for creating the three Microsoft VBA signature profiles on Windows.</summary>
    public sealed class WordMacroProjectSigningOptions : WordMacroProjectSignatureValidationOptions {
        /// <summary>Gets or sets the path to SignTool.exe. When omitted, SignTool.exe is resolved through PATH.</summary>
        public string? SignToolPath { get; set; }

        /// <summary>Gets or sets the maximum duration of one external tool invocation. Defaults to two minutes.</summary>
        public TimeSpan ToolTimeout { get; set; } = TimeSpan.FromMinutes(2);

        /// <summary>Gets or sets the maximum captured output characters per tool invocation. Defaults to 64 KiB.</summary>
        public int MaxToolOutputCharacters { get; set; } = 64 * 1024;

        /// <summary>Gets or sets the certificate store name. Defaults to Personal (My).</summary>
        public StoreName StoreName { get; set; } = StoreName.My;

        /// <summary>Gets or sets the certificate store location. Defaults to CurrentUser.</summary>
        public StoreLocation StoreLocation { get; set; } = StoreLocation.CurrentUser;

        /// <summary>Gets or sets the directory containing Microsoft's offclearsig.exe. Required for signing.</summary>
        public string? OfficeSipsDirectory { get; set; }

        /// <summary>Gets or sets an optional RFC 3161 timestamp-authority URL.</summary>
        public Uri? TimestampAuthorityUrl { get; set; }

        /// <summary>Gets or sets how existing OPC package signatures are handled. The default blocks VBA signing.</summary>
        public WordSignedDocumentSavePolicy ExistingPackageSignaturePolicy { get; set; } = WordSignedDocumentSavePolicy.Block;
    }

    /// <summary>Content-binding, certificate, revocation, and timestamp validation of a VBA signature.</summary>
    public sealed class WordMacroProjectSignatureValidationResult {
        internal WordMacroProjectSignatureValidationResult(
            WordMacroProjectSignatureInfo signatureInfo,
            bool isSupported,
            WordSignatureValidationState contentBindingStatus,
            IReadOnlyList<WordMacroProjectSignatureFinding> findings) {
            SignatureInfo = signatureInfo;
            IsSupported = isSupported;
            ContentBindingStatus = contentBindingStatus;
            Findings = findings;
        }

        /// <summary>Gets cross-platform VBA signature metadata and CMS policy evidence.</summary>
        public WordMacroProjectSignatureInfo SignatureInfo { get; }

        /// <summary>Gets whether native VBA content-binding validation is supported in the current environment.</summary>
        public bool IsSupported { get; }

        /// <summary>Gets Microsoft Office SIP validation of the selected VBA signature against project content.</summary>
        public WordSignatureValidationState ContentBindingStatus { get; }

        /// <summary>Gets stable aggregate validation diagnostics.</summary>
        public IReadOnlyList<WordMacroProjectSignatureFinding> Findings { get; }

        /// <summary>Gets whether content binding, signer trust, requested revocation policy, and timestamp policy passed.</summary>
        public bool IsValidUnderPolicy { get; internal set; }
    }

    /// <summary>Result of an atomic VBA macro-project signing attempt.</summary>
    public sealed class WordMacroProjectSigningResult {
        internal WordMacroProjectSigningResult(
            string filePath,
            bool isSupported,
            bool succeeded,
            bool macroProjectPreserved,
            WordMacroProjectSignatureValidationResult? validationResult,
            IReadOnlyList<WordMacroProjectSignatureFinding> findings) {
            FilePath = filePath;
            IsSupported = isSupported;
            Succeeded = succeeded;
            MacroProjectPreserved = macroProjectPreserved;
            ValidationResult = validationResult;
            Findings = findings;
        }

        /// <summary>Gets the signed document path.</summary>
        public string FilePath { get; }

        /// <summary>Gets whether VBA signing is supported in the current environment.</summary>
        public bool IsSupported { get; }

        /// <summary>Gets whether the staged document was signed, validated, and atomically committed.</summary>
        public bool Succeeded { get; }

        /// <summary>Gets whether the encoded VBA project bytes were unchanged by signing.</summary>
        public bool MacroProjectPreserved { get; }

        /// <summary>Gets validation readback from the completed staging file.</summary>
        public WordMacroProjectSignatureValidationResult? ValidationResult { get; }

        /// <summary>Gets stable signing and validation diagnostics.</summary>
        public IReadOnlyList<WordMacroProjectSignatureFinding> Findings { get; }
    }

    /// <summary>Raised when VBA macro-project signing could not be completed and proven.</summary>
    public sealed class WordMacroProjectSigningException : InvalidOperationException {
        internal WordMacroProjectSigningException(WordMacroProjectSigningResult result)
            : base(result.Findings.LastOrDefault()?.Message ?? "VBA macro-project signing failed.") {
            Result = result;
        }

        /// <summary>Gets the structured failed signing result.</summary>
        public WordMacroProjectSigningResult Result { get; }
    }
}
