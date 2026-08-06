using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Core.Internal;
using OfficeIMO.Security;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        /// <summary>
        /// Inspects package-level digital-signature metadata without copying the encoded package or validating digests and cryptographic trust.
        /// Use <see cref="ValidateSignatures(IOfficeSecurityProvider, WordSignatureValidationOptions?)"/> for transform-aware digest and trust validation.
        /// </summary>
        public WordSignatureInfo InspectSignatures() {
            using var snapshot = new MemoryStream();
            using (_wordprocessingDocument.Clone(snapshot)) { }
            OfficeIMO.Security.OfficePackageSignatureInfo shared = OfficePackageSignatureService.Inspect(
                snapshot.ToArray(), new OfficePackageSignatureInspectionOptions { VerifyDigests = false });
            return CreateWordSignatureInfo(shared,
                _wordprocessingDocument.DigitalSignatureOriginPart != null,
                ApplicationProperties.DigitalSignature != null);
        }

        private static WordSignatureInfo CreateWordSignatureInfo(
            OfficeIMO.Security.OfficePackageSignatureInfo shared,
            bool hasLiveOrigin,
            bool hasLiveApplicationMetadata) {
            const string signatureContentType =
                "application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml";
            WordSignaturePartInfo[] parts = shared.SignatureParts.Select(part =>
                new WordSignaturePartInfo(
                    part.Uri,
                    signatureContentType,
                    relationshipId: null,
                    part.Length,
                    part.SignatureMethodAlgorithm,
                    part.SignedReferences.Select(reference => reference.DigestMethodAlgorithm)
                        .Where(algorithm => !string.IsNullOrWhiteSpace(algorithm))
                        .Select(algorithm => algorithm!)
                        .Distinct(StringComparer.Ordinal)
                        .ToArray(),
                    part.SignedReferences.Select(reference => new WordSignatureReferenceInfo(
                        reference.Uri,
                        reference.DigestMethodAlgorithm,
                        reference.DigestValue,
                        reference.IsPackagePartReference,
                        reference.TargetPartUri,
                        reference.TargetPartExists,
                        reference.TransformAlgorithms,
                        MapSharedSignatureState(reference.DigestVerificationStatus),
                        reference.DigestVerificationDetail)).ToArray(),
                    part.Timestamps.Select(timestamp => new WordSignatureTimestampInfo(
                        timestamp.Kind, timestamp.Value, timestamp.Format)).ToArray(),
                    part.X509SubjectNames,
                    part.ParseError,
                    part.IsReachableFromOrigin
                        ? Array.Empty<string>()
                        : new[] { "The XML signature part is not reachable through the package signature-origin relationship chain." }))
                .ToArray();
            var details = new List<string>(shared.Findings);
            if (!string.IsNullOrWhiteSpace(shared.OriginPartUri)) {
                details.Add("Digital signature origin part: " + shared.OriginPartUri);
            }
            details.AddRange(shared.SignatureParts.Select(part => "XML signature part: " + part.Uri));
            if (shared.HasApplicationSignatureMetadata || hasLiveApplicationMetadata) {
                details.Add("Extended application properties advertise digital signature metadata.");
            }
            return new WordSignatureInfo(
                shared.HasDigitalSignatureOriginPart || hasLiveOrigin,
                shared.OriginPartUri,
                originRelationshipId: null,
                shared.HasApplicationSignatureMetadata || hasLiveApplicationMetadata,
                parts,
                shared.Findings,
                details);
        }

        private static WordSignatureValidationState MapSharedSignatureState(
            OfficePackageSignatureValidationState state) => state switch {
                OfficePackageSignatureValidationState.Passed => WordSignatureValidationState.Passed,
                OfficePackageSignatureValidationState.Failed => WordSignatureValidationState.Failed,
                OfficePackageSignatureValidationState.Unsupported => WordSignatureValidationState.Unsupported,
                _ => WordSignatureValidationState.NotChecked
            };

        /// <summary>
        /// Validates package structure, transform-aware OPC digests, XML signature math, signer trust,
        /// revocation, and timestamps through an explicitly supplied security provider.
        /// </summary>
        /// <param name="securityProvider">Cryptographic provider supplied by the caller, typically <c>OfficeSecurityProvider.Default</c>.</param>
        /// <param name="options">Trust, revocation, timestamp, and resource policy.</param>
        public WordSignatureValidationReport ValidateSignatures(
            IOfficeSecurityProvider securityProvider,
            WordSignatureValidationOptions? options = null) {
            if (securityProvider == null) throw new ArgumentNullException(nameof(securityProvider));
            options ??= new WordSignatureValidationOptions();
            OfficePackageSignatureValidator.ValidateOptions(options);
            var originPart = _wordprocessingDocument.DigitalSignatureOriginPart;
            bool hasApplicationSignatureMetadata = ApplicationProperties.DigitalSignature != null;
            if (originPart == null || !originPart.XmlSignatureParts.Any()) {
                WordSignatureInfo unsignedInfo = WordSignatureInspector.Inspect(
                    _wordprocessingDocument,
                    originPart,
                    hasApplicationSignatureMetadata,
                    packageBytes: null,
                    maxPackageParts: options.MaxPackageParts,
                    maxPartBytes: options.MaxPartBytes,
                    maxSignedReferences: options.MaxSignedReferences,
                    maxTotalDigestBytes: options.MaxTotalDigestBytes,
                    maxSignatureBytes: options.MaxSignatureBytes,
                    maxCertificates: options.MaxCertificates,
                    maxCertificateBytes: options.MaxCertificateBytes,
                    maxTotalCertificateBytes: options.MaxTotalCertificateBytes,
                    verifyDigests: false);
                return WordSignatureValidationReport.From(unsignedInfo);
            }
            if (originPart != null && originPart.XmlSignatureParts.Skip(options.MaxSignatureParts).Any()) {
                var boundedInfo = new WordSignatureInfo(
                    hasDigitalSignatureOriginPart: true,
                    originPart!.Uri.ToString(),
                    originRelationshipId: null,
                    hasApplicationSignatureMetadata,
                    Array.Empty<WordSignaturePartInfo>(),
                    Array.Empty<string>(),
                    new[] { "Digital-signature inspection stopped before parsing because the signature-part count exceeded policy." });
                return WordSignatureValidationReport.WithValidationFailure(
                    WordSignatureValidationReport.From(boundedInfo),
                    "SignatureResourceLimitExceeded",
                    "The package contains more than " + options.MaxSignatureParts + " XML signature parts.");
            }
            byte[] packageBytes;
            try {
                packageBytes = CreateSignatureValidationSnapshot(options);
            } catch (SignatureValidationSnapshotResourceException exception) {
                var boundedInfo = new WordSignatureInfo(
                    hasDigitalSignatureOriginPart: true,
                    originPart!.Uri.ToString(),
                    originRelationshipId: null,
                    hasApplicationSignatureMetadata,
                    Array.Empty<WordSignaturePartInfo>(),
                    Array.Empty<string>(),
                    new[] { exception.Message });
                return WordSignatureValidationReport.WithValidationFailure(
                    WordSignatureValidationReport.From(boundedInfo),
                    "SignatureResourceLimitExceeded",
                    exception.Message);
            } catch (InvalidDataException exception) {
                WordSignatureInfo boundedInfo = WordSignatureInspector.Inspect(
                    _wordprocessingDocument,
                    originPart,
                    hasApplicationSignatureMetadata,
                    packageBytes: null,
                    maxPackageParts: options.MaxPackageParts,
                    maxPartBytes: options.MaxPartBytes,
                    maxSignedReferences: options.MaxSignedReferences,
                    maxTotalDigestBytes: options.MaxTotalDigestBytes,
                    maxSignatureBytes: options.MaxSignatureBytes,
                    maxCertificates: options.MaxCertificates,
                    maxCertificateBytes: options.MaxCertificateBytes,
                    maxTotalCertificateBytes: options.MaxTotalCertificateBytes,
                    verifyDigests: false);
                return WordSignatureValidationReport.WithValidationFailure(
                    WordSignatureValidationReport.From(boundedInfo),
                    "PackageByteLimitExceeded",
                    exception.Message);
            }

            using var validationStream = new MemoryStream(packageBytes, writable: false);
            using WordprocessingDocument validationPackage = WordprocessingDocument.Open(validationStream, false);
            DigitalSignatureOriginPart? validationOriginPart = validationPackage.DigitalSignatureOriginPart;
            bool validationHasApplicationSignatureMetadata = validationPackage.ExtendedFilePropertiesPart?.Properties?.DigitalSignature != null;
            WordSignatureInfo signatureInfo = WordSignatureInspector.Inspect(
                validationPackage,
                validationOriginPart,
                validationHasApplicationSignatureMetadata,
                packageBytes,
                options.MaxPackageParts,
                options.MaxPartBytes,
                options.MaxSignedReferences,
                options.MaxTotalDigestBytes,
                options.MaxSignatureBytes,
                options.MaxCertificates,
                options.MaxCertificateBytes,
                options.MaxTotalCertificateBytes,
                securityProvider: securityProvider);
            WordSignatureValidationReport structural = WordSignatureValidationReport.From(signatureInfo);
            if (signatureInfo.InspectionResourceLimitExceeded) {
                return WordSignatureValidationReport.WithValidationFailure(
                    structural,
                    "SignatureResourceLimitExceeded",
                    signatureInfo.UnsupportedDetails.FirstOrDefault() ??
                    "Digital-signature inspection stopped at the configured package resource limit.");
            }
            if (!signatureInfo.HasSignatures || validationOriginPart == null) {
                return structural;
            }

            try {
                IReadOnlyList<WordSignaturePartValidationResult> signatures = OfficePackageSignatureValidator.Validate(
                    validationOriginPart,
                    packageBytes,
                    signatureInfo,
                    securityProvider,
                    options);
                return WordSignatureValidationReport.WithCryptographicValidation(structural, signatures);
            } catch (InvalidDataException exception) {
                return WordSignatureValidationReport.WithValidationFailure(
                    structural,
                    "SignatureResourceLimitExceeded",
                    exception.Message);
            }
        }

        /// <summary>
        /// Signs a saved DOCX package using the cross-platform OPC XML-signature engine and throws when signing cannot be completed and cryptographically verified.
        /// </summary>
        /// <param name="filePath">Path to the DOCX package to sign.</param>
        /// <param name="securityProvider">Cryptographic provider supplied by the optional OfficeIMO.Security package.</param>
        /// <param name="certificate">Certificate with a private key used for signing.</param>
        /// <param name="options">Optional package-signing settings.</param>
        /// <returns>A signing result with structural, cryptographic, digest, and certificate-policy validation readback.</returns>
        public static WordPackageSigningResult SignPackage(
            string filePath,
            IOfficeSecurityProvider securityProvider,
            X509Certificate2 certificate,
            WordPackageSigningOptions? options = null) {
            return RequireSuccessfulSigningReadback(TrySignPackage(filePath, securityProvider, certificate, options));
        }

        /// <summary>
        /// Resolves a signing certificate by thumbprint from the certificate store, signs a saved DOCX package, and throws when signing cannot be completed and cryptographically verified.
        /// </summary>
        /// <param name="filePath">Path to the DOCX package to sign.</param>
        /// <param name="securityProvider">Cryptographic provider supplied by the optional OfficeIMO.Security package.</param>
        /// <param name="certificateThumbprint">Certificate thumbprint to locate.</param>
        /// <param name="certificateOptions">Optional certificate-store lookup settings.</param>
        /// <param name="signingOptions">Optional package-signing settings.</param>
        /// <returns>A signing result with structural, cryptographic, digest, and certificate-policy validation readback.</returns>
        public static WordPackageSigningResult SignPackage(
            string filePath,
            IOfficeSecurityProvider securityProvider,
            string certificateThumbprint,
            WordPackageCertificateStoreOptions? certificateOptions = null,
            WordPackageSigningOptions? signingOptions = null) {
            return RequireSuccessfulSigningReadback(TrySignPackage(
                filePath,
                securityProvider,
                certificateThumbprint,
                certificateOptions,
                signingOptions));
        }

        private static WordPackageSigningResult RequireSuccessfulSigningReadback(WordPackageSigningResult result) {
            if (!result.CreatedSignatureReadbackSucceeded) throw new WordPackageSigningException(result);
            return result;
        }

        /// <summary>
        /// Attempts to sign a saved DOCX package and returns a report instead of throwing for unsupported platforms or signing failures.
        /// </summary>
        /// <param name="filePath">Path to the DOCX package to sign.</param>
        /// <param name="securityProvider">Cryptographic provider supplied by the optional OfficeIMO.Security package.</param>
        /// <param name="certificate">Certificate with a private key used for signing.</param>
        /// <param name="options">Optional package-signing settings.</param>
        /// <returns>A signing result with details and validation readback when available.</returns>
        public static WordPackageSigningResult TrySignPackage(
            string filePath,
            IOfficeSecurityProvider securityProvider,
            X509Certificate2 certificate,
            WordPackageSigningOptions? options = null) {
            if (securityProvider == null) throw new ArgumentNullException(nameof(securityProvider));
            WordPackageSigningOptions effectiveOptions = options ?? new WordPackageSigningOptions();
            WordSignatureValidationReport? validationReport = null;
            string? createdSignaturePartUri = null;
            OfficePackageSigningOptions packageOptions = effectiveOptions.ToPackageOptions();
            packageOptions.ValidateBeforeCommit = (stagingPath, signaturePartUri, signatureCount) => {
                createdSignaturePartUri = signaturePartUri;
                using WordDocument document = Load(stagingPath, CreateSigningReadbackLoadOptions(effectiveOptions));
                validationReport = document.ValidateSignatures(
                    securityProvider,
                    CreateSigningReadbackOptions(effectiveOptions, signatureCount));
                WordSignaturePartValidationResult? createdSignature = validationReport.Signatures.FirstOrDefault(signature =>
                    string.Equals(signature.SignaturePart.Uri, signaturePartUri, StringComparison.OrdinalIgnoreCase));
                if (WordPackageSigningResult.IsCreatedSignatureValidationSuccessful(createdSignature)) return null;

                string detail = validationReport.Findings.FirstOrDefault()
                    ?? "The created signature was missing or failed cryptographic or package-reference digest validation.";
                return "The created package signature failed validation readback before atomic commit: " + detail;
            };
            OfficePackageSigningResult packageResult = OfficePackageSignatureWriter.Sign(
                filePath,
                certificate,
                securityProvider,
                packageOptions);

            return new WordPackageSigningResult(packageResult, validationReport, createdSignaturePartUri);
        }

        internal static WordLoadOptions CreateSigningReadbackLoadOptions(
            WordPackageSigningOptions signingOptions) => new WordLoadOptions {
                AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly,
                MaxInputBytes = signingOptions.MaxPackageBytes
            };

        internal static WordSignatureValidationOptions CreateSigningReadbackOptions(
            WordPackageSigningOptions signingOptions,
            int signatureCount) {
            return new WordSignatureValidationOptions {
                MaxSignatureParts = Math.Max(32, signatureCount),
                MaxPackageBytes = signingOptions.MaxPackageBytes,
                MaxPackageParts = signingOptions.MaxPackageParts,
                MaxPartBytes = signingOptions.MaxPartBytes,
                MaxSignedReferences = signingOptions.MaxSignedReferences,
                MaxTotalDigestBytes = signingOptions.MaxTotalDigestBytes,
                MaxSignatureBytes = signingOptions.MaxSignatureBytes,
                MaxCertificates = signingOptions.MaxCertificates,
                MaxCertificateBytes = signingOptions.MaxCertificateBytes,
                MaxTotalCertificateBytes = Math.Max(
                    WordSignatureValidationOptions.DefaultMaxTotalCertificateBytes,
                    MultiplyLimit(signingOptions.MaxTotalCertificateBytes, Math.Max(1, signatureCount)))
            };
        }

        private static long MultiplyLimit(long value, int multiplier) =>
            value > long.MaxValue / multiplier ? long.MaxValue : value * multiplier;

        /// <summary>
        /// Attempts to resolve a signing certificate by thumbprint from the certificate store and sign a saved DOCX package.
        /// </summary>
        /// <param name="filePath">Path to the DOCX package to sign.</param>
        /// <param name="securityProvider">Cryptographic provider supplied by the optional OfficeIMO.Security package.</param>
        /// <param name="certificateThumbprint">Certificate thumbprint to locate.</param>
        /// <param name="certificateOptions">Optional certificate-store lookup settings.</param>
        /// <param name="signingOptions">Optional package-signing settings.</param>
        /// <returns>A signing result with details and validation readback when available.</returns>
        public static WordPackageSigningResult TrySignPackage(
            string filePath,
            IOfficeSecurityProvider securityProvider,
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
                return TrySignPackage(fullPath, securityProvider, certificate!, signingOptions);
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
