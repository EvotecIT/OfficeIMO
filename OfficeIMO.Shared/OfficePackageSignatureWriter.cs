#nullable enable
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography.X509Certificates;

#if NET472
using System.IO.Packaging;
#endif

namespace OfficeIMO.Shared {
    /// <summary>
    /// Options for signing an Open Packaging Convention package.
    /// </summary>
    internal sealed class OfficePackageSigningOptions {
        internal const string Sha256HashAlgorithm = "http://www.w3.org/2001/04/xmlenc#sha256";

        /// <summary>Gets or sets explicit package-part URIs to sign. When null, all existing non-signature package parts are signed.</summary>
        public IReadOnlyCollection<string>? PartUris { get; set; }

        /// <summary>Gets or sets whether package-level relationships are signed.</summary>
        public bool IncludePackageRelationships { get; set; } = true;

        /// <summary>Gets or sets whether part-level relationships are signed.</summary>
        public bool IncludePartRelationships { get; set; } = true;

        /// <summary>Gets or sets the package-signature hash algorithm URI.</summary>
        public string HashAlgorithm { get; set; } = Sha256HashAlgorithm;

        /// <summary>Gets or sets an optional signature id.</summary>
        public string? SignatureId { get; set; }
    }

    /// <summary>
    /// Result of an attempted Open Packaging Convention package-signing operation.
    /// </summary>
    internal sealed class OfficePackageSigningResult {
        internal OfficePackageSigningResult(
            string filePath,
            bool isSupported,
            bool succeeded,
            int signedPartCount,
            int signedRelationshipSelectorCount,
            int signatureCount,
            string? signaturePartUri,
            IReadOnlyList<string> details) {
            FilePath = filePath;
            IsSupported = isSupported;
            Succeeded = succeeded;
            SignedPartCount = signedPartCount;
            SignedRelationshipSelectorCount = signedRelationshipSelectorCount;
            SignatureCount = signatureCount;
            SignaturePartUri = signaturePartUri;
            Details = details;
        }

        public string FilePath { get; }

        public bool IsSupported { get; }

        public bool Succeeded { get; }

        public int SignedPartCount { get; }

        public int SignedRelationshipSelectorCount { get; }

        public int SignatureCount { get; }

        public string? SignaturePartUri { get; }

        public IReadOnlyList<string> Details { get; }
    }

    /// <summary>
    /// Signs Open Packaging Convention packages when the current target framework exposes package-signing primitives.
    /// </summary>
    internal static class OfficePackageSignatureWriter {
        internal static OfficePackageSigningResult Sign(string filePath, X509Certificate2 certificate, OfficePackageSigningOptions? options = null) {
            options ??= new OfficePackageSigningOptions();

            if (string.IsNullOrWhiteSpace(filePath)) {
                return Failed(filePath ?? string.Empty, true, "A package path is required.");
            }

            string fullPath = Path.GetFullPath(filePath);
            if (!File.Exists(fullPath)) {
                return Failed(fullPath, true, "The package file does not exist.");
            }

            if (certificate == null) {
                return Failed(fullPath, true, "A signing certificate is required.");
            }

            if (!certificate.HasPrivateKey) {
                return Failed(fullPath, true, "The signing certificate must include a private key.");
            }

#if NET472
            try {
                using Package package = Package.Open(fullPath, FileMode.Open, FileAccess.ReadWrite);
                List<Uri> partUris = ResolvePartUris(package, options, out IReadOnlyList<string> missingPartUris);
                List<PackageRelationshipSelector> relationshipSelectors = ResolveRelationshipSelectors(package, options).ToList();

                if (missingPartUris.Count > 0) {
                    return Failed(fullPath, true, "Requested signing part(s) were not found: " + string.Join(", ", missingPartUris) + ".");
                }

                if (partUris.Count == 0) {
                    return Failed(fullPath, true, "No package parts were selected for signing.");
                }

                var manager = new PackageDigitalSignatureManager(package) {
                    CertificateOption = CertificateEmbeddingOption.InCertificatePart,
                };

                if (!string.IsNullOrWhiteSpace(options.HashAlgorithm)) {
                    manager.HashAlgorithm = options.HashAlgorithm;
                }

                PackageDigitalSignature signature = string.IsNullOrWhiteSpace(options.SignatureId)
                    ? manager.Sign(partUris, certificate, relationshipSelectors)
                    : manager.Sign(partUris, certificate, relationshipSelectors, options.SignatureId);

                var details = new List<string> {
                    "Package signature was created with " + partUris.Count.ToString(System.Globalization.CultureInfo.InvariantCulture) + " signed package part(s).",
                    "Signed relationship selector count: " + relationshipSelectors.Count.ToString(System.Globalization.CultureInfo.InvariantCulture) + ".",
                    "Signature count after signing: " + manager.Signatures.Count.ToString(System.Globalization.CultureInfo.InvariantCulture) + "."
                };

                return new OfficePackageSigningResult(
                    fullPath,
                    isSupported: true,
                    succeeded: true,
                    signedPartCount: partUris.Count,
                    signedRelationshipSelectorCount: relationshipSelectors.Count,
                    signatureCount: manager.Signatures.Count,
                    signaturePartUri: signature.SignaturePart.Uri.ToString(),
                    details: details.ToArray());
            } catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException || ex is InvalidOperationException || ex is ArgumentException || ex is System.Security.Cryptography.CryptographicException) {
                return Failed(fullPath, true, "Package signing failed: " + ex.Message);
            }
#else
            string[] details = {
                "Package signing requires System.IO.Packaging.PackageDigitalSignatureManager, which is not available to this OfficeIMO.Word target framework.",
                "Use the Windows .NET Framework signing adapter or an external package-signing service/tool, then inspect the signed DOCX with WordDocument.ValidateSignatures()."
            };
            return new OfficePackageSigningResult(
                fullPath,
                isSupported: false,
                succeeded: false,
                signedPartCount: 0,
                signedRelationshipSelectorCount: 0,
                signatureCount: 0,
                signaturePartUri: null,
                details: details);
#endif
        }

        private static OfficePackageSigningResult Failed(string filePath, bool isSupported, string detail) {
            return new OfficePackageSigningResult(
                filePath,
                isSupported,
                succeeded: false,
                signedPartCount: 0,
                signedRelationshipSelectorCount: 0,
                signatureCount: 0,
                signaturePartUri: null,
                details: new[] { detail });
        }

#if NET472
        private static List<Uri> ResolvePartUris(Package package, OfficePackageSigningOptions options, out IReadOnlyList<string> missingPartUris) {
            HashSet<string>? requestedUris = options.PartUris == null
                ? null
                : new HashSet<string>(options.PartUris.Select(NormalizePartUri), StringComparer.OrdinalIgnoreCase);

            var resolvedUris = new List<Uri>();
            var foundRequestedUris = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (PackagePart part in package.GetParts().OrderBy(part => part.Uri.ToString(), StringComparer.OrdinalIgnoreCase)) {
                string normalizedPartUri = NormalizePartUri(part.Uri.ToString());
                if (normalizedPartUri.StartsWith("/_xmlsignatures/", StringComparison.OrdinalIgnoreCase)) {
                    continue;
                }

                if (IsRelationshipPart(part, normalizedPartUri)) {
                    continue;
                }

                if (requestedUris == null || requestedUris.Contains(normalizedPartUri)) {
                    resolvedUris.Add(part.Uri);
                    foundRequestedUris.Add(normalizedPartUri);
                }
            }

            missingPartUris = requestedUris == null
                ? Array.Empty<string>()
                : requestedUris.Where(uri => !foundRequestedUris.Contains(uri)).OrderBy(uri => uri, StringComparer.OrdinalIgnoreCase).ToArray();
            return resolvedUris;
        }

        private static IEnumerable<PackageRelationshipSelector> ResolveRelationshipSelectors(Package package, OfficePackageSigningOptions options) {
            if (options.IncludePackageRelationships) {
                foreach (PackageRelationship relationship in package.GetRelationships().OrderBy(relationship => relationship.Id, StringComparer.OrdinalIgnoreCase)) {
                    yield return new PackageRelationshipSelector(new Uri("/", UriKind.Relative), PackageRelationshipSelectorType.Id, relationship.Id);
                }
            }

            if (!options.IncludePartRelationships) {
                yield break;
            }

            foreach (PackagePart part in package.GetParts().OrderBy(part => part.Uri.ToString(), StringComparer.OrdinalIgnoreCase)) {
                string normalizedPartUri = NormalizePartUri(part.Uri.ToString());
                if (normalizedPartUri.StartsWith("/_xmlsignatures/", StringComparison.OrdinalIgnoreCase)) {
                    continue;
                }

                if (IsRelationshipPart(part, normalizedPartUri)) {
                    continue;
                }

                foreach (PackageRelationship relationship in part.GetRelationships().OrderBy(relationship => relationship.Id, StringComparer.OrdinalIgnoreCase)) {
                    yield return new PackageRelationshipSelector(part.Uri, PackageRelationshipSelectorType.Id, relationship.Id);
                }
            }
        }

        private static string NormalizePartUri(string partUri) {
            if (string.IsNullOrWhiteSpace(partUri)) {
                return "/";
            }

            string normalized = partUri.Trim().Replace('\\', '/');
            return normalized.StartsWith("/", StringComparison.Ordinal) ? normalized : "/" + normalized;
        }

        private static bool IsRelationshipPart(PackagePart part, string normalizedPartUri) {
            return normalizedPartUri.EndsWith(".rels", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(part.ContentType, "application/vnd.openxmlformats-package.relationships+xml", StringComparison.OrdinalIgnoreCase);
        }
#endif
    }
}
