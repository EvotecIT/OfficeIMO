using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing.Internal;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        /// <summary>
        /// Inspects package-level digital-signature metadata without copying the encoded package or validating digests and cryptographic trust.
        /// Use <see cref="ValidateSignatures()"/> for transform-aware digest and trust validation.
        /// </summary>
        public WordSignatureInfo InspectSignatures() {
            var originPart = _wordprocessingDocument.DigitalSignatureOriginPart;
            return WordSignatureInspector.Inspect(
                _wordprocessingDocument,
                originPart,
                ApplicationProperties.DigitalSignature != null,
                packageBytes: null,
                verifyDigests: false);
        }

        /// <summary>
        /// Validates package structure, transform-aware OPC digests, XML signature math, signer trust,
        /// revocation under the default no-network policy, and embedded RFC 3161 timestamp tokens.
        /// </summary>
        public WordSignatureValidationReport ValidateSignatures() {
            return ValidateSignatures(new WordSignatureValidationOptions());
        }

        /// <summary>
        /// Validates package structure, transform-aware OPC digests, XML signature math, signer trust,
        /// revocation, and embedded RFC 3161 timestamp tokens under caller policy.
        /// </summary>
        /// <param name="options">Trust, revocation, timestamp, and resource policy.</param>
        public WordSignatureValidationReport ValidateSignatures(WordSignatureValidationOptions options) {
            if (options == null) throw new ArgumentNullException(nameof(options));
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
                options.MaxTotalCertificateBytes);
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
        /// <param name="certificate">Certificate with a private key used for signing.</param>
        /// <param name="options">Optional package-signing settings.</param>
        /// <returns>A signing result with structural, cryptographic, digest, and certificate-policy validation readback.</returns>
        public static WordPackageSigningResult SignPackage(string filePath, X509Certificate2 certificate, WordPackageSigningOptions? options = null) {
            return RequireSuccessfulSigningReadback(TrySignPackage(filePath, certificate, options));
        }

        /// <summary>
        /// Resolves a signing certificate by thumbprint from the certificate store, signs a saved DOCX package, and throws when signing cannot be completed and cryptographically verified.
        /// </summary>
        /// <param name="filePath">Path to the DOCX package to sign.</param>
        /// <param name="certificateThumbprint">Certificate thumbprint to locate.</param>
        /// <param name="certificateOptions">Optional certificate-store lookup settings.</param>
        /// <param name="signingOptions">Optional package-signing settings.</param>
        /// <returns>A signing result with structural, cryptographic, digest, and certificate-policy validation readback.</returns>
        public static WordPackageSigningResult SignPackage(
            string filePath,
            string certificateThumbprint,
            WordPackageCertificateStoreOptions? certificateOptions = null,
            WordPackageSigningOptions? signingOptions = null) {
            return RequireSuccessfulSigningReadback(TrySignPackage(filePath, certificateThumbprint, certificateOptions, signingOptions));
        }

        private static WordPackageSigningResult RequireSuccessfulSigningReadback(WordPackageSigningResult result) {
            if (!result.CreatedSignatureReadbackSucceeded) throw new WordPackageSigningException(result);
            return result;
        }

        /// <summary>
        /// Attempts to sign a saved DOCX package and returns a report instead of throwing for unsupported platforms or signing failures.
        /// </summary>
        /// <param name="filePath">Path to the DOCX package to sign.</param>
        /// <param name="certificate">Certificate with a private key used for signing.</param>
        /// <param name="options">Optional package-signing settings.</param>
        /// <returns>A signing result with details and validation readback when available.</returns>
        public static WordPackageSigningResult TrySignPackage(string filePath, X509Certificate2 certificate, WordPackageSigningOptions? options = null) {
            WordPackageSigningOptions effectiveOptions = options ?? new WordPackageSigningOptions();
            OfficePackageSigningResult packageResult = OfficePackageSignatureWriter.Sign(filePath, certificate, effectiveOptions.ToPackageOptions());
            WordSignatureValidationReport? validationReport = null;

            if (packageResult.Succeeded) {
                using WordDocument document = Load(filePath, CreateSigningReadbackLoadOptions(effectiveOptions));
                validationReport = document.ValidateSignatures(CreateSigningReadbackOptions(
                    effectiveOptions,
                    packageResult.SignatureCount));
            }

            return new WordPackageSigningResult(packageResult, validationReport);
        }

        internal static WordLoadOptions CreateSigningReadbackLoadOptions(
            WordPackageSigningOptions signingOptions) => new WordLoadOptions {
                AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly,
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

        private byte[] CreateSignatureValidationSnapshot(WordSignatureValidationOptions options) {
            if (_ownedPackageStream == null) {
                throw new InvalidDataException("The current OPC package has no encoded package stream available for validation.");
            }
            if (_wordprocessingDocument.FileOpenAccess == FileAccess.Read) {
                if (_ownedPackageStream.Length > options.MaxPackageBytes) {
                    throw new InvalidDataException("The current OPC package exceeds the " + options.MaxPackageBytes + " byte validation limit.");
                }
                return _ownedPackageStream.ToArray();
            }

            byte[] encodedPackage = _ownedPackageStream.ToArray();
            using var snapshot = new MemoryStream(encodedPackage.Length);
            snapshot.Write(encodedPackage, 0, encodedPackage.Length);
            snapshot.Position = 0;
            using (WordprocessingDocument snapshotPackage = WordprocessingDocument.Open(
                snapshot,
                true,
                new OpenSettings { AutoSave = false })) {
                Dictionary<Uri, OpenXmlPart> snapshotParts = EnumerateSignatureSnapshotParts(
                        snapshotPackage,
                        options.MaxPackageParts)
                    .ToDictionary(part => part.Uri);
                foreach (OpenXmlPart sourcePart in EnumerateSignatureSnapshotParts(
                    _wordprocessingDocument,
                    options.MaxPackageParts)) {
                    OpenXmlPartRootElement? sourceRoot = sourcePart.IsRootElementLoaded
                        ? sourcePart.RootElement
                        : null;
                    if (sourceRoot == null ||
                        !snapshotParts.TryGetValue(sourcePart.Uri, out OpenXmlPart? snapshotPart)) {
                        continue;
                    }
                    using (Stream input = snapshotPart.GetStream(FileMode.Open, FileAccess.Read)) {
                        if (input.Length > options.MaxPartBytes) {
                            throw new SignatureValidationSnapshotResourceException(
                                "The loaded package part " + sourcePart.Uri + " exceeds the " +
                                options.MaxPartBytes + " byte validation limit.");
                        }
                    }
                    OpenXmlPartRootElement? snapshotRoot = snapshotPart.RootElement;
                    if (snapshotRoot != null && AreSignatureSnapshotRootsEquivalent(sourceRoot, snapshotRoot)) {
                        continue;
                    }
                    using Stream output = snapshotPart.GetStream(FileMode.Create, FileAccess.Write);
                    ((OpenXmlPartRootElement)sourceRoot.CloneNode(true)).Save(output);
                }
            }
            if (snapshot.Length > options.MaxPackageBytes) {
                throw new InvalidDataException("The current OPC package exceeds the " + options.MaxPackageBytes + " byte validation limit.");
            }
            return snapshot.ToArray();
        }

        private static IEnumerable<OpenXmlPart> EnumerateSignatureSnapshotParts(
            OpenXmlPartContainer container,
            int maxPackageParts) {
            var pending = new Stack<OpenXmlPart>(container.Parts.Select(pair => pair.OpenXmlPart));
            var visited = new HashSet<Uri>();
            while (pending.Count > 0) {
                OpenXmlPart part = pending.Pop();
                if (!visited.Add(part.Uri)) continue;
                if (visited.Count > maxPackageParts) {
                    throw new SignatureValidationSnapshotResourceException(
                        "The OPC package contains more than " + maxPackageParts + " parts during validation snapshot creation.");
                }
                yield return part;
                foreach (IdPartPair child in part.Parts) pending.Push(child.OpenXmlPart);
            }
        }

        private static bool AreSignatureSnapshotRootsEquivalent(OpenXmlElement source, OpenXmlElement snapshot) {
            var pending = new Stack<(OpenXmlElement Source, OpenXmlElement Snapshot)>();
            pending.Push((source, snapshot));
            while (pending.Count > 0) {
                (OpenXmlElement sourceElement, OpenXmlElement snapshotElement) = pending.Pop();
                if (!string.Equals(sourceElement.LocalName, snapshotElement.LocalName, StringComparison.Ordinal) ||
                    !string.Equals(sourceElement.NamespaceUri, snapshotElement.NamespaceUri, StringComparison.Ordinal)) {
                    return false;
                }
                IList<OpenXmlAttribute> sourceAttributes = sourceElement.GetAttributes();
                IList<OpenXmlAttribute> snapshotAttributes = snapshotElement.GetAttributes();
                if (sourceAttributes.Count != snapshotAttributes.Count || sourceAttributes.Any(sourceAttribute =>
                    !snapshotAttributes.Any(snapshotAttribute =>
                        string.Equals(sourceAttribute.LocalName, snapshotAttribute.LocalName, StringComparison.Ordinal) &&
                        string.Equals(sourceAttribute.NamespaceUri, snapshotAttribute.NamespaceUri, StringComparison.Ordinal) &&
                        string.Equals(sourceAttribute.Value, snapshotAttribute.Value, StringComparison.Ordinal)))) {
                    return false;
                }
                List<KeyValuePair<string, string>> sourceNamespaces = sourceElement.NamespaceDeclarations.ToList();
                List<KeyValuePair<string, string>> snapshotNamespaces = snapshotElement.NamespaceDeclarations.ToList();
                if (sourceNamespaces.Count != snapshotNamespaces.Count || sourceNamespaces.Any(sourceNamespace =>
                    !snapshotNamespaces.Any(snapshotNamespace =>
                        string.Equals(sourceNamespace.Key, snapshotNamespace.Key, StringComparison.Ordinal) &&
                        string.Equals(sourceNamespace.Value, snapshotNamespace.Value, StringComparison.Ordinal)))) {
                    return false;
                }
                if (sourceElement.ChildElements.Count != snapshotElement.ChildElements.Count) return false;
                if (sourceElement.ChildElements.Count == 0 && !string.Equals(
                    sourceElement.InnerText,
                    snapshotElement.InnerText,
                    StringComparison.Ordinal)) {
                    return false;
                }
                for (int index = sourceElement.ChildElements.Count - 1; index >= 0; index--) {
                    pending.Push((sourceElement.ChildElements[index], snapshotElement.ChildElements[index]));
                }
            }
            return true;
        }

        private sealed class SignatureValidationSnapshotResourceException : Exception {
            internal SignatureValidationSnapshotResourceException(string message) : base(message) { }
        }

        /// <summary>
        /// Attempts to resolve a signing certificate by thumbprint from the certificate store and sign a saved DOCX package.
        /// </summary>
        /// <param name="filePath">Path to the DOCX package to sign.</param>
        /// <param name="certificateThumbprint">Certificate thumbprint to locate.</param>
        /// <param name="certificateOptions">Optional certificate-store lookup settings.</param>
        /// <param name="signingOptions">Optional package-signing settings.</param>
        /// <returns>A signing result with details and validation readback when available.</returns>
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
