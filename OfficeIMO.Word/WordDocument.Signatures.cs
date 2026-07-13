using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using OfficeIMO.Drawing.Internal;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        /// <summary>
        /// Inspects package-level digital-signature metadata without validating cryptographic trust.
        /// </summary>
        public WordSignatureInfo InspectSignatures() {
            return WordSignatureInspector.Inspect(
                _wordprocessingDocument,
                _wordprocessingDocument.DigitalSignatureOriginPart,
                ApplicationProperties.DigitalSignature != null);
        }

        /// <summary>
        /// Validates signature package structure and reports unsupported cryptographic validation boundaries.
        /// </summary>
        public WordSignatureValidationReport ValidateSignatures() {
            return WordSignatureValidationReport.From(InspectSignatures());
        }

        /// <summary>
        /// Signs a saved DOCX package using the platform package-signing adapter and throws when signing cannot be completed and structurally verified.
        /// </summary>
        /// <param name="filePath">Path to the DOCX package to sign.</param>
        /// <param name="certificate">Certificate with a private key used for signing.</param>
        /// <param name="options">Optional package-signing settings.</param>
        /// <returns>A signing result with structural validation readback.</returns>
        public static WordPackageSigningResult SignPackage(string filePath, X509Certificate2 certificate, WordPackageSigningOptions? options = null) {
            WordPackageSigningResult result = TrySignPackage(filePath, certificate, options);
            if (!result.Succeeded || result.ValidationReport?.IsStructurallyValid != true) {
                throw new WordPackageSigningException(result);
            }

            return result;
        }

        /// <summary>
        /// Resolves a signing certificate by thumbprint from the certificate store, signs a saved DOCX package, and throws when signing cannot be completed and structurally verified.
        /// </summary>
        /// <param name="filePath">Path to the DOCX package to sign.</param>
        /// <param name="certificateThumbprint">Certificate thumbprint to locate.</param>
        /// <param name="certificateOptions">Optional certificate-store lookup settings.</param>
        /// <param name="signingOptions">Optional package-signing settings.</param>
        /// <returns>A signing result with structural validation readback.</returns>
        public static WordPackageSigningResult SignPackage(
            string filePath,
            string certificateThumbprint,
            WordPackageCertificateStoreOptions? certificateOptions = null,
            WordPackageSigningOptions? signingOptions = null) {
            WordPackageSigningResult result = TrySignPackage(filePath, certificateThumbprint, certificateOptions, signingOptions);
            if (!result.Succeeded || result.ValidationReport?.IsStructurallyValid != true) {
                throw new WordPackageSigningException(result);
            }

            return result;
        }

        /// <summary>
        /// Attempts to sign a saved DOCX package and returns a report instead of throwing for unsupported platforms or signing failures.
        /// </summary>
        /// <param name="filePath">Path to the DOCX package to sign.</param>
        /// <param name="certificate">Certificate with a private key used for signing.</param>
        /// <param name="options">Optional package-signing settings.</param>
        /// <returns>A signing result with details and structural validation readback when available.</returns>
        public static WordPackageSigningResult TrySignPackage(string filePath, X509Certificate2 certificate, WordPackageSigningOptions? options = null) {
            OfficePackageSigningResult packageResult = OfficePackageSignatureWriter.Sign(filePath, certificate, (options ?? new WordPackageSigningOptions()).ToPackageOptions());
            WordSignatureValidationReport? validationReport = null;

            if (packageResult.Succeeded) {
                using WordDocument document = Load(filePath, new WordLoadOptions {
                    AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly
                });
                validationReport = document.ValidateSignatures();
            }

            return new WordPackageSigningResult(packageResult, validationReport);
        }

        /// <summary>
        /// Attempts to resolve a signing certificate by thumbprint from the certificate store and sign a saved DOCX package.
        /// </summary>
        /// <param name="filePath">Path to the DOCX package to sign.</param>
        /// <param name="certificateThumbprint">Certificate thumbprint to locate.</param>
        /// <param name="certificateOptions">Optional certificate-store lookup settings.</param>
        /// <param name="signingOptions">Optional package-signing settings.</param>
        /// <returns>A signing result with details and structural validation readback when available.</returns>
        public static WordPackageSigningResult TrySignPackage(
            string filePath,
            string certificateThumbprint,
            WordPackageCertificateStoreOptions? certificateOptions = null,
            WordPackageSigningOptions? signingOptions = null) {
            string fullPath = string.IsNullOrWhiteSpace(filePath)
                ? filePath ?? string.Empty
                : Path.GetFullPath(filePath);

            if (!TryResolveSigningCertificate(certificateThumbprint, certificateOptions, out X509Certificate2? certificate, out string detail)) {
                return WordPackageSigningResult.Failed(fullPath, isSupported: true, new[] { detail });
            }

            using (certificate) {
                return TrySignPackage(fullPath, certificate!, signingOptions);
            }
        }

        private static bool TryResolveSigningCertificate(
            string certificateThumbprint,
            WordPackageCertificateStoreOptions? options,
            out X509Certificate2? certificate,
            out string detail) {
            certificate = null;
            options ??= new WordPackageCertificateStoreOptions();

            if (!TryNormalizeCertificateThumbprint(certificateThumbprint, out string normalizedThumbprint, out string validationDetail)) {
                detail = validationDetail;
                return false;
            }

            if (string.IsNullOrWhiteSpace(normalizedThumbprint)) {
                detail = "A certificate thumbprint is required.";
                return false;
            }

            try {
                using var store = new X509Store(options.StoreName, options.StoreLocation);
                store.Open(OpenFlags.ReadOnly | OpenFlags.OpenExistingOnly);
                X509Certificate2Collection matches = store.Certificates.Find(
                    X509FindType.FindByThumbprint,
                    normalizedThumbprint,
                    !options.IncludeInvalidCertificates);

                X509Certificate2? match = matches
                    .OfType<X509Certificate2>()
                    .FirstOrDefault(item => !options.RequirePrivateKey || item.HasPrivateKey);
                if (match == null) {
                    detail = "Certificate thumbprint " + normalizedThumbprint + " was not found in "
                        + options.StoreLocation + "\\" + options.StoreName
                        + (options.RequirePrivateKey ? " with an accessible private key." : ".");
                    return false;
                }

                certificate = new X509Certificate2(match);
                detail = "Resolved signing certificate from " + options.StoreLocation + "\\" + options.StoreName + ".";
                return true;
            } catch (Exception ex) when (ex is CryptographicException || ex is PlatformNotSupportedException || ex is UnauthorizedAccessException) {
                detail = "Certificate store lookup failed for " + options.StoreLocation + "\\" + options.StoreName + ": " + ex.Message;
                return false;
            }
        }

        private static bool TryNormalizeCertificateThumbprint(string? thumbprint, out string normalizedThumbprint, out string detail) {
            normalizedThumbprint = string.Empty;
            detail = string.Empty;
            if (string.IsNullOrWhiteSpace(thumbprint)) {
                return true;
            }

            string value = thumbprint!;
            var chars = new List<char>(value.Length);
            foreach (char character in value) {
                if (Uri.IsHexDigit(character)) {
                    chars.Add(char.ToUpperInvariant(character));
                } else if (char.IsWhiteSpace(character) || character == ':' || character == '-') {
                    continue;
                } else {
                    detail = "Certificate thumbprint contains invalid character '" + character + "'.";
                    return false;
                }
            }

            normalizedThumbprint = new string(chars.ToArray());
            return true;
        }
    }
}
