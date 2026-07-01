using System.Security.Cryptography.X509Certificates;
using OfficeIMO.Shared;

namespace OfficeIMO.Word {
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
    /// Options for signing a DOCX package through the platform package-signing adapter.
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

        internal OfficePackageSigningOptions ToPackageOptions() {
            return new OfficePackageSigningOptions {
                PartUris = PartUris,
                IncludePackageRelationships = IncludePackageRelationships,
                IncludePartRelationships = IncludePartRelationships,
                HashAlgorithm = HashAlgorithm,
                SignatureId = SignatureId
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
            if (validationReport != null && !validationReport.IsStructurallyValid) {
                details.AddRange(validationReport.Findings);
            }

            FilePath = packageResult.FilePath;
            IsSupported = packageResult.IsSupported;
            Succeeded = packageResult.Succeeded;
            SignedPartCount = packageResult.SignedPartCount;
            SignedRelationshipSelectorCount = packageResult.SignedRelationshipSelectorCount;
            SignatureCount = packageResult.SignatureCount;
            SignaturePartUri = packageResult.SignaturePartUri;
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

        /// <summary>Gets structural validation readback for the signed package when signing succeeded.</summary>
        public WordSignatureValidationReport? ValidationReport { get; }

        /// <summary>Gets deterministic signing details or failure reasons.</summary>
        public IReadOnlyList<string> Details { get; }

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
