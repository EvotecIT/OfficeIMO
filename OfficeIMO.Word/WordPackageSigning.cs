using System.Security.Cryptography.X509Certificates;
using System.Runtime.InteropServices;
using OfficeIMO.Drawing.Internal;

namespace OfficeIMO.Word {
    /// <summary>Identifies a distinct Word signing surface.</summary>
    public enum WordSigningCapabilityKind {
        /// <summary>OPC XML signature over package parts and relationships.</summary>
        OpcPackage,
        /// <summary>VBA project signature embedded in a macro-enabled package.</summary>
        VbaMacroProject
    }

    /// <summary>Describes availability of one signing surface without conflating package and VBA signatures.</summary>
    public sealed class WordSigningCapability {
        internal WordSigningCapability(WordSigningCapabilityKind kind, bool isSupported, string message) {
            Kind = kind;
            IsSupported = isSupported;
            Message = message;
        }

        /// <summary>Gets the signing surface.</summary>
        public WordSigningCapabilityKind Kind { get; }
        /// <summary>Gets whether OfficeIMO can create this signature type.</summary>
        public bool IsSupported { get; }
        /// <summary>Gets the precise capability boundary.</summary>
        public string Message { get; }
    }

    /// <summary>Exposes package and macro-project signing as separate capabilities.</summary>
    public static class WordSigningCapabilities {
        /// <summary>Cross-platform OPC package-signature creation and validation.</summary>
        public static WordSigningCapability Package { get; } = new WordSigningCapability(
            WordSigningCapabilityKind.OpcPackage,
            true,
            "Cross-platform OPC XML package signing and validation are supported.");

        /// <summary>VBA project signing capability, which is independent of OPC package signing.</summary>
        public static WordSigningCapability MacroProject { get; } = new WordSigningCapability(
            WordSigningCapabilityKind.VbaMacroProject,
            RuntimeInformation.IsOSPlatform(OSPlatform.Windows),
            RuntimeInformation.IsOSPlatform(OSPlatform.Windows)
                ? "VBA macro-project signing is available through Microsoft OfficeSips and SignTool as a separate Windows-native capability; operational availability is reported per request."
                : "VBA signature parts can be inspected cross-platform as a separate capability, but content-binding validation requires Microsoft's registered Office SIP on Windows and creation additionally requires OfficeSips tooling plus SignTool.");
    }

    /// <summary>
    /// Options for resolving a signing certificate from the local certificate store.
    /// </summary>
    public sealed class WordPackageCertificateStoreOptions {
        /// <summary>
        /// Gets or sets the certificate store name to inspect.
        /// </summary>
        public StoreName StoreName { get; set; } = StoreName.My;

        /// <summary>
        /// Gets or sets the certificate store location to inspect.
        /// </summary>
        public StoreLocation StoreLocation { get; set; } = StoreLocation.CurrentUser;

        /// <summary>
        /// Gets or sets whether the resolved certificate must include a private key.
        /// </summary>
        public bool RequirePrivateKey { get; set; } = true;

        /// <summary>
        /// Gets or sets whether invalid or untrusted certificates are included during thumbprint lookup.
        /// </summary>
        public bool IncludeInvalidCertificates { get; set; } = true;
    }

    /// <summary>
    /// Options for signing a DOCX package through the cross-platform OPC XML-signature engine.
    /// </summary>
    public sealed class WordPackageSigningOptions {
        /// <summary>
        /// Gets the SHA-256 XML DSig hash algorithm URI used by default.
        /// </summary>
        public const string Sha256HashAlgorithm = OfficePackageSigningOptions.Sha256HashAlgorithm;

        /// <summary>
        /// Gets or sets explicit package-part URIs to sign. When null, all existing non-signature package parts are signed.
        /// </summary>
        public IReadOnlyCollection<string>? PartUris { get; set; }

        /// <summary>
        /// Gets or sets whether package-level relationships are included in the signature.
        /// </summary>
        public bool IncludePackageRelationships { get; set; } = true;

        /// <summary>
        /// Gets or sets whether relationships owned by individual package parts are included in the signature.
        /// </summary>
        public bool IncludePartRelationships { get; set; } = true;

        /// <summary>
        /// Gets or sets the XML DSig hash algorithm URI.
        /// </summary>
        public string HashAlgorithm { get; set; } = Sha256HashAlgorithm;

        /// <summary>
        /// Gets or sets an optional signature id.
        /// </summary>
        public string? SignatureId { get; set; }

        /// <summary>
        /// Gets or sets the claimed OPC signing time. Current UTC time is used when omitted.
        /// This is signed metadata, not an RFC 3161 timestamp-authority token.
        /// </summary>
        public DateTimeOffset? SigningTime { get; set; }

        /// <summary>Gets or sets optional intermediate or root certificates embedded with the signer certificate.</summary>
        public IReadOnlyCollection<X509Certificate2>? AdditionalCertificates { get; set; }

        /// <summary>Gets or sets the maximum number of ZIP entries accepted while signing. Defaults to 10,000.</summary>
        public int MaxPackageParts { get; set; } = 10000;

        /// <summary>Gets or sets the maximum encoded package bytes accepted while signing. Defaults to 512 MiB.</summary>
        public long MaxPackageBytes { get; set; } = 512L * 1024 * 1024;

        /// <summary>Gets or sets the maximum uncompressed bytes read from one signed part. Defaults to 256 MiB.</summary>
        public long MaxPartBytes { get; set; } = 256L * 1024 * 1024;

        /// <summary>Gets or sets the maximum aggregate uncompressed package-part bytes digested while signing. Defaults to 512 MiB.</summary>
        public long MaxTotalDigestBytes { get; set; } = 512L * 1024 * 1024;

        /// <summary>Gets or sets the maximum authenticated XML signature references created per signature. Defaults to 4,096.</summary>
        public int MaxSignedReferences { get; set; } = 4096;

        /// <summary>Gets or sets the maximum encoded bytes for the generated XML signature part. Defaults to 16 MiB.</summary>
        public long MaxSignatureBytes { get; set; } = 16L * 1024 * 1024;

        /// <summary>Gets or sets the maximum signer and additional certificates embedded in the created signature. Defaults to 64.</summary>
        public int MaxCertificates { get; set; } = 64;

        /// <summary>Gets or sets the maximum encoded bytes for one embedded certificate. Defaults to 4 MiB.</summary>
        public long MaxCertificateBytes { get; set; } = 4L * 1024 * 1024;

        /// <summary>Gets or sets the maximum aggregate encoded certificate bytes embedded in the created signature. Defaults to 64 MiB.</summary>
        public long MaxTotalCertificateBytes { get; set; } = 64L * 1024 * 1024;

        internal OfficePackageSigningOptions ToPackageOptions() {
            return new OfficePackageSigningOptions {
                PartUris = PartUris,
                IncludePackageRelationships = IncludePackageRelationships,
                IncludePartRelationships = IncludePartRelationships,
                HashAlgorithm = HashAlgorithm,
                SignatureId = SignatureId,
                SigningTime = SigningTime,
                AdditionalCertificates = AdditionalCertificates,
                MaxPackageParts = MaxPackageParts,
                MaxPackageBytes = MaxPackageBytes,
                MaxPartBytes = MaxPartBytes,
                MaxTotalDigestBytes = MaxTotalDigestBytes,
                MaxSignedReferences = MaxSignedReferences,
                MaxSignatureBytes = MaxSignatureBytes,
                MaxCertificates = MaxCertificates,
                MaxCertificateBytes = MaxCertificateBytes,
                MaxTotalCertificateBytes = MaxTotalCertificateBytes
            };
        }
    }

    /// <summary>
    /// Describes the result of a DOCX package-signing attempt.
    /// </summary>
    public sealed class WordPackageSigningResult {
        internal WordPackageSigningResult(
            OfficePackageSigningResult packageResult,
            WordSignatureValidationReport? validationReport) {
            var details = new List<string>(packageResult.Details);
            if (validationReport != null && !validationReport.IsValidUnderPolicy) {
                details.AddRange(validationReport.Findings);
            }

            FilePath = packageResult.FilePath;
            IsSupported = packageResult.IsSupported;
            Succeeded = packageResult.Succeeded;
            SignedPartCount = packageResult.SignedPartCount;
            SignedRelationshipSelectorCount = packageResult.SignedRelationshipSelectorCount;
            SignatureCount = packageResult.SignatureCount;
            SignaturePartUri = packageResult.SignaturePartUri;
            CreatedSignatureValidation = validationReport?.Signatures.FirstOrDefault(signature =>
                string.Equals(signature.SignaturePart.Uri, packageResult.SignaturePartUri, StringComparison.OrdinalIgnoreCase));
            Details = details;
            ValidationReport = validationReport;
        }

        /// <summary>Gets the signed package path.</summary>
        public string FilePath { get; }

        /// <summary>Gets whether the current target framework supports package signing.</summary>
        public bool IsSupported { get; }

        /// <summary>Gets whether a package signature was created.</summary>
        public bool Succeeded { get; }

        /// <summary>Gets the number of package parts selected for signing.</summary>
        public int SignedPartCount { get; }

        /// <summary>Gets the number of package relationship selectors included in the signature.</summary>
        public int SignedRelationshipSelectorCount { get; }

        /// <summary>Gets the signature count reported by the package-signing adapter after signing.</summary>
        public int SignatureCount { get; }

        /// <summary>Gets the generated signature part URI when signing succeeded.</summary>
        public string? SignaturePartUri { get; }

        /// <summary>Gets validation readback for the signature created by this signing operation.</summary>
        public WordSignaturePartValidationResult? CreatedSignatureValidation { get; }

        /// <summary>Gets aggregate cryptographic, digest, certificate, revocation, and timestamp validation readback for every signature in the package.</summary>
        public WordSignatureValidationReport? ValidationReport { get; }

        /// <summary>Gets deterministic signing details or failure reasons.</summary>
        public IReadOnlyList<string> Details { get; }

        internal bool CreatedSignatureReadbackSucceeded {
            get {
                WordSignaturePartValidationResult? validation = CreatedSignatureValidation;
                WordSignaturePartInfo? signaturePart = validation?.SignaturePart;
                if (!Succeeded ||
                    validation?.CryptographicStatus != WordSignatureValidationState.Passed ||
                    signaturePart == null ||
                    signaturePart.HasParseError ||
                    string.IsNullOrWhiteSpace(signaturePart.SignatureMethodAlgorithm) ||
                    signaturePart.SignedReferences.Count == 0) {
                    return false;
                }

                return signaturePart.SignedReferences.All(reference =>
                    reference.IsPackagePartReference &&
                    reference.TargetPartExists == true &&
                    !string.IsNullOrWhiteSpace(reference.DigestMethodAlgorithm) &&
                    reference.HasDigestValue &&
                    reference.DigestVerificationStatus == WordSignatureValidationState.Passed);
            }
        }

        internal static WordPackageSigningResult Failed(string filePath, bool isSupported, IReadOnlyList<string> details) {
            return new WordPackageSigningResult(
                filePath,
                isSupported,
                details);
        }

        private WordPackageSigningResult(string filePath, bool isSupported, IReadOnlyList<string> details) {
            FilePath = filePath;
            IsSupported = isSupported;
            Succeeded = false;
            SignedPartCount = 0;
            SignedRelationshipSelectorCount = 0;
            SignatureCount = 0;
            SignaturePartUri = null;
            CreatedSignatureValidation = null;
            Details = details.ToArray();
            ValidationReport = null;
        }
    }

    /// <summary>
    /// Raised when DOCX package signing was requested but could not be completed and proven.
    /// </summary>
    public sealed class WordPackageSigningException : InvalidOperationException {
        internal WordPackageSigningException(WordPackageSigningResult result)
            : base(CreateMessage(result)) {
            Result = result;
        }

        /// <summary>
        /// Gets the failed signing result.
        /// </summary>
        public WordPackageSigningResult Result { get; }

        private static string CreateMessage(WordPackageSigningResult result) {
            string detail = result.Details.Count == 0 ? "No signing detail was provided." : result.Details[0];
            return "DOCX package signing failed for '" + result.FilePath + "'. " + detail;
        }
    }
}
