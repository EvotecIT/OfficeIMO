#nullable enable
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Security;
using System.Collections;
using System.Diagnostics.CodeAnalysis;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Security.Cryptography.Xml;
using System.Xml;

namespace OfficeIMO.Word {
    /// <summary>Cross-platform XML DSig, X.509, revocation, and RFC 3161 validator for OPC signatures.</summary>
    internal static partial class OfficePackageSignatureValidator {
        private static readonly HashSet<string> XadesNamespaces = new(StringComparer.Ordinal) {
            "http://uri.etsi.org/01903/v1.1.1#",
            "http://uri.etsi.org/01903/v1.2.2#",
            "http://uri.etsi.org/01903/v1.3.2#",
            "http://uri.etsi.org/01903/v1.4.1#"
        };
        private static readonly HashSet<string> SupportedSignedInfoReferenceTransforms = new(StringComparer.Ordinal) {
            SignedXml.XmlDsigC14NTransformUrl,
            SignedXml.XmlDsigC14NWithCommentsTransformUrl,
            SignedXml.XmlDsigExcC14NTransformUrl,
            SignedXml.XmlDsigExcC14NWithCommentsTransformUrl,
            SignedXml.XmlDsigEnvelopedSignatureTransformUrl
        };
        private static readonly HashSet<string> SupportedSignedInfoCanonicalizationMethods = new(StringComparer.Ordinal) {
            SignedXml.XmlDsigC14NTransformUrl,
            SignedXml.XmlDsigC14NWithCommentsTransformUrl,
            SignedXml.XmlDsigExcC14NTransformUrl,
            SignedXml.XmlDsigExcC14NWithCommentsTransformUrl
        };

        internal static IReadOnlyList<WordSignaturePartValidationResult> Validate(
            DigitalSignatureOriginPart? originPart,
            byte[] packageBytes,
            WordSignatureInfo signatureInfo,
            WordSignatureValidationOptions options) {
            if (originPart == null || signatureInfo.SignatureParts.Count == 0) {
                return Array.Empty<WordSignaturePartValidationResult>();
            }
            ValidateOptions(options);
            if (signatureInfo.SignatureParts.Count > options.MaxSignatureParts) {
                throw new InvalidDataException("The package contains more than " + options.MaxSignatureParts + " XML signature parts.");
            }

            if (packageBytes.LongLength > options.MaxPackageBytes) {
                throw new InvalidDataException("The OPC package exceeds the " + options.MaxPackageBytes + " byte validation limit.");
            }
            using var archive = new OfficePackageSignatureArchive(packageBytes, options.MaxPackageParts);
            var certificateByteBudget = new OfficePackageCertificateByteBudget(options.MaxTotalCertificateBytes);
            var timestampBudget = new OfficePackageTimestampValidationBudget(
                options.MaxTimestampTokens,
                options.MaxTimestampBytes);
            Dictionary<string, XmlSignaturePart> packageParts = originPart.XmlSignatureParts
                .ToDictionary(part => OfficePackageSignatureArchive.NormalizePartUri(part.Uri.ToString()), StringComparer.OrdinalIgnoreCase);
            var results = new List<WordSignaturePartValidationResult>(signatureInfo.SignatureParts.Count);
            foreach (WordSignaturePartInfo signaturePartInfo in signatureInfo.SignatureParts) {
                if (!packageParts.TryGetValue(
                    OfficePackageSignatureArchive.NormalizePartUri(signaturePartInfo.Uri),
                    out XmlSignaturePart? signaturePart)) {
                    results.Add(FailedMissingPart(signaturePartInfo));
                    continue;
                }
                results.Add(ValidateSignaturePart(
                    signaturePart,
                    signaturePartInfo,
                    archive,
                    options,
                    certificateByteBudget,
                    timestampBudget));
            }
            return results;
        }

        private static WordSignaturePartValidationResult ValidateSignaturePart(
            XmlSignaturePart signaturePart,
            WordSignaturePartInfo signaturePartInfo,
            OfficePackageSignatureArchive archive,
            WordSignatureValidationOptions options,
            OfficePackageCertificateByteBudget certificateByteBudget,
            OfficePackageTimestampValidationBudget timestampBudget) {
            var findings = new List<WordSignatureValidationFinding>();
            var timestampResults = new List<Rfc3161TimestampVerificationResult>();
            var certificates = new List<X509Certificate2>();
            bool revocationCheckRequired = options.CertificateValidation.RevocationMode != X509RevocationMode.NoCheck;
            try {
                byte[] signatureBytes = archive.ReadPart(signaturePartInfo.Uri, options.MaxSignatureBytes);
                XmlDocument document = LoadXml(signatureBytes, options.MaxSignatureBytes);
                XmlElement? signatureElement = document.DocumentElement;
                if (signatureElement == null ||
                    signatureElement.LocalName != "Signature" ||
                    signatureElement.NamespaceURI != SignedXml.XmlDsigNamespaceUrl) {
                    return FailedMalformed(signaturePartInfo, "The signature part does not contain an XML DSig Signature root element.");
                }
                EnsureSignedInfoReferenceCountWithinLimit(signatureElement, options.MaxSignedReferences);

                certificates.AddRange(ReadCertificates(
                    signaturePart,
                    document,
                    options.MaxCertificates,
                    options.MaxCertificateBytes,
                    certificateByteBudget,
                    findings));
                XmlElement? signatureValue = ReadSignatureValue(document, findings, signaturePartInfo.Uri);
                WordSignatureValidationState cryptographicStatus;
                IReadOnlyList<X509Certificate2> matchingSigners = Array.Empty<X509Certificate2>();

                if (!options.ValidateCryptographicSignature) {
                    cryptographicStatus = WordSignatureValidationState.NotChecked;
                    findings.Add(Finding("CryptographicValidationDisabled", cryptographicStatus,
                        "XML DSig signature-value validation was disabled by caller policy.", signaturePartInfo.Uri));
                } else if (certificates.Count == 0) {
                    cryptographicStatus = WordSignatureValidationState.Unsupported;
                    findings.Add(Finding("SignerCertificateMissing", cryptographicStatus,
                        "No embedded or related X.509 signer certificate was found.", signaturePartInfo.Uri));
                } else {
                    cryptographicStatus = ValidateSignedXml(
                        document,
                        signatureElement,
                        certificates,
                        options.MaxTotalDigestBytes,
                        signaturePartInfo.Uri,
                        findings,
                        out matchingSigners);
                }

                WordSignatureValidationState timestampStatus;
                IReadOnlyList<XmlElement> timestampTokens = GetXadesTimestampTokens(signatureElement);
                try {
                    if (options.ValidateTimestamps && signatureValue != null) {
                        ValidateTimestampTokens(
                            timestampTokens,
                            signatureValue,
                            options,
                            signaturePartInfo.Uri,
                            timestampBudget,
                            timestampResults,
                            findings);
                    }

                    timestampStatus = ResolveTimestampStatus(
                        document,
                        timestampTokens.Count,
                        options,
                        timestampResults,
                        signaturePartInfo.Uri,
                        findings);
                } catch (InvalidDataException exception) {
                    timestampStatus = WordSignatureValidationState.Failed;
                    findings.Add(Finding(
                        "TimestampResourceLimitExceeded",
                        timestampStatus,
                        "Timestamp validation exceeds the configured resource limits: " + exception.Message,
                        signaturePartInfo.Uri));
                }

                CertificateValidationResult? certificateValidation = null;
                WordSignatureValidationState certificateStatus;
                WordSignatureValidationState revocationStatus;
                if (matchingSigners.Count == 0) {
                    certificateStatus = certificates.Count == 0
                        ? WordSignatureValidationState.NotPresent
                        : WordSignatureValidationState.NotChecked;
                    revocationStatus = certificateStatus;
                } else {
                    CertificateValidationOptions signerOptions = ResolveSignerCertificateValidation(
                        options.CertificateValidation,
                        timestampResults);
                    CertificateTrustValidationResult trust = SelectSignerTrust(
                        matchingSigners,
                        certificates,
                        signerOptions,
                        revocationCheckRequired);
                    certificateValidation = trust.Validation;
                    certificateStatus = MapStatus(trust.Validation.ChainStatus);
                    revocationStatus = MapStatus(trust.Validation.RevocationStatus);
                    foreach (SecurityFinding finding in trust.Findings) {
                        findings.Add(Finding(
                            finding.Code,
                            finding.Severity == SecurityFindingSeverity.Error
                                ? WordSignatureValidationState.Failed
                                : certificateStatus,
                            finding.Message,
                            signaturePartInfo.Uri));
                    }
                }

                return new WordSignaturePartValidationResult(
                    signaturePartInfo,
                    cryptographicStatus,
                    certificateStatus,
                    revocationStatus,
                    revocationCheckRequired,
                    timestampStatus,
                    certificateValidation,
                    timestampResults.ToArray(),
                    findings.ToArray());
            } catch (InvalidDataException exception) {
                findings.Add(Finding("SignatureResourceLimitExceeded", WordSignatureValidationState.Failed,
                    "The XML signature exceeds the configured validation resource limits: " + exception.Message, signaturePartInfo.Uri));
                return new WordSignaturePartValidationResult(
                    signaturePartInfo,
                    WordSignatureValidationState.Failed,
                    WordSignatureValidationState.NotChecked,
                    WordSignatureValidationState.NotChecked,
                    revocationCheckRequired,
                    WordSignatureValidationState.NotChecked,
                    null,
                    timestampResults.ToArray(),
                    findings.ToArray());
            } catch (Exception exception) when (IsValidationException(exception)) {
                findings.Add(Finding("SignatureValidationFailed", WordSignatureValidationState.Failed,
                    "The XML signature could not be validated: " + exception.Message, signaturePartInfo.Uri));
                return new WordSignaturePartValidationResult(
                    signaturePartInfo,
                    WordSignatureValidationState.Failed,
                    WordSignatureValidationState.NotChecked,
                    WordSignatureValidationState.NotChecked,
                    revocationCheckRequired,
                    WordSignatureValidationState.NotChecked,
                    null,
                    timestampResults.ToArray(),
                    findings.ToArray());
            } finally {
                foreach (X509Certificate2 certificate in certificates) certificate.Dispose();
            }
        }

        private static void EnsureSignedInfoReferenceCountWithinLimit(
            XmlElement signatureElement,
            int maxSignedReferences) {
            XmlElement? signedInfo = signatureElement.ChildNodes
                .OfType<XmlElement>()
                .FirstOrDefault(element =>
                    element.LocalName == "SignedInfo" &&
                    element.NamespaceURI == SignedXml.XmlDsigNamespaceUrl);
            int referenceCount = signedInfo?.ChildNodes
                .OfType<XmlElement>()
                .Count(element =>
                    element.LocalName == "Reference" &&
                    element.NamespaceURI == SignedXml.XmlDsigNamespaceUrl) ?? 0;
            if (referenceCount > maxSignedReferences) {
                throw new InvalidDataException(
                    "The XML signature contains more than " + maxSignedReferences + " SignedInfo references.");
            }
        }

        [UnconditionalSuppressMessage("Trimming", "IL2026", Justification = "OPC signature validation accepts only the explicitly handled XML DSig algorithms; their implementations are referenced directly by OfficeIMO and preserved for trimmed applications.")]
        [UnconditionalSuppressMessage("AOT", "IL3050", Justification = "OfficeIMO does not enable the XSLT XML DSig transform; OPC signature validation is limited to the statically supported transform set and does not compile dynamic XSLT code.")]
        private static WordSignatureValidationState ValidateSignedXml(
            XmlDocument document,
            XmlElement signatureElement,
            IReadOnlyList<X509Certificate2> certificates,
            long maxTotalDigestBytes,
            string signaturePartUri,
            List<WordSignatureValidationFinding> findings,
            out IReadOnlyList<X509Certificate2> matchingSigners) {
            matchingSigners = Array.Empty<X509Certificate2>();
            if (!HasSupportedSignedInfoCanonicalizationMethod(signatureElement, out string? unsupportedCanonicalization)) {
                findings.Add(Finding(
                    "UnsupportedSignedInfoCanonicalizationMethod",
                    WordSignatureValidationState.Unsupported,
                    "SignedInfo canonicalization method '" + unsupportedCanonicalization + "' is outside the supported canonicalization profile.",
                    signaturePartUri));
                return WordSignatureValidationState.Unsupported;
            }
            if (!HasOnlySupportedSignedInfoReferenceTransforms(signatureElement, out string? unsupportedTransform)) {
                findings.Add(Finding(
                    "UnsupportedSignedInfoTransform",
                    WordSignatureValidationState.Unsupported,
                    "SignedInfo reference transform '" + unsupportedTransform + "' is outside the supported canonicalization and enveloped-signature profile.",
                    signaturePartUri));
                return WordSignatureValidationState.Unsupported;
            }
            var signedXml = new SignedXml(document) { Resolver = null! };
            try {
                signedXml.LoadXml(signatureElement);
            } catch (CryptographicException exception) {
                findings.Add(Finding("XmlSignatureMalformed", WordSignatureValidationState.Failed,
                    "The XML DSig structure is invalid: " + exception.Message, signaturePartUri));
                return WordSignatureValidationState.Failed;
            }

            if (!HasOnlyLocalSignedInfoReferences(signedXml, out string? unsupportedUri)) {
                findings.Add(Finding("ExternalSignedInfoReference", WordSignatureValidationState.Unsupported,
                    "SignedInfo reference '" + unsupportedUri + "' is not a local fragment and was not dereferenced.",
                    signaturePartUri,
                    unsupportedUri));
                return WordSignatureValidationState.Unsupported;
            }
            EnsureLocalSignedInfoDigestWorkWithinLimit(
                document,
                signedXml,
                certificates.Count,
                maxTotalDigestBytes);

            var matches = new List<X509Certificate2>();
            foreach (X509Certificate2 candidate in certificates) {
                try {
                    using AsymmetricAlgorithm? publicKey = GetPublicKey(candidate);
                    if (publicKey == null || !signedXml.CheckSignature(publicKey)) continue;
                    matches.Add(candidate);
                } catch (CryptographicException) {
                    // Try the remaining embedded or related certificate candidates.
                }
            }

            if (matches.Count > 0) {
                matchingSigners = matches;
                findings.Add(Finding("XmlSignatureValid", WordSignatureValidationState.Passed,
                    "XML DSig signature-value and signed-object validation passed.", signaturePartUri));
                return WordSignatureValidationState.Passed;
            }

            findings.Add(Finding("XmlSignatureInvalid", WordSignatureValidationState.Failed,
                "XML DSig signature-value or signed-object validation failed for every supplied certificate.", signaturePartUri));
            return WordSignatureValidationState.Failed;
        }

        private static CertificateTrustValidationResult SelectSignerTrust(
            IReadOnlyList<X509Certificate2> matchingSigners,
            IReadOnlyList<X509Certificate2> certificates,
            CertificateValidationOptions options,
            bool revocationCheckRequired) {
            CertificateTrustValidationResult? fallback = null;
            foreach (X509Certificate2 signer in matchingSigners) {
                CertificateTrustValidationResult trust = CertificateValidator.Validate(
                    signer,
                    certificates.Where(certificate => !ReferenceEquals(certificate, signer)),
                    options,
                    CertificateValidationPurpose.DocumentSigning);
                fallback ??= trust;
                if (IsSignerTrustAccepted(trust.Validation, revocationCheckRequired)) return trust;
            }
            return fallback!;
        }

        private static bool IsSignerTrustAccepted(
            CertificateValidationResult validation,
            bool revocationCheckRequired) =>
            validation.ChainStatus == SecurityValidationStatus.Valid &&
            validation.RevocationStatus != SecurityValidationStatus.Invalid &&
            (!revocationCheckRequired || validation.RevocationStatus == SecurityValidationStatus.Valid);

        private static bool HasOnlySupportedSignedInfoReferenceTransforms(
            XmlElement signatureElement,
            out string? unsupportedTransform) {
            unsupportedTransform = null;
            XmlElement? signedInfo = signatureElement.ChildNodes
                .OfType<XmlElement>()
                .FirstOrDefault(element =>
                    element.LocalName == "SignedInfo" &&
                    element.NamespaceURI == SignedXml.XmlDsigNamespaceUrl);
            if (signedInfo == null) return true;

            foreach (XmlElement reference in signedInfo.ChildNodes
                         .OfType<XmlElement>()
                         .Where(element =>
                             element.LocalName == "Reference" &&
                             element.NamespaceURI == SignedXml.XmlDsigNamespaceUrl)) {
                XmlElement? transforms = reference.ChildNodes
                    .OfType<XmlElement>()
                    .FirstOrDefault(element =>
                        element.LocalName == "Transforms" &&
                        element.NamespaceURI == SignedXml.XmlDsigNamespaceUrl);
                if (transforms == null) continue;
                foreach (XmlElement transform in transforms.ChildNodes
                             .OfType<XmlElement>()
                             .Where(element =>
                                 element.LocalName == "Transform" &&
                                 element.NamespaceURI == SignedXml.XmlDsigNamespaceUrl)) {
                    string algorithm = transform.GetAttribute("Algorithm").Trim();
                    if (SupportedSignedInfoReferenceTransforms.Contains(algorithm)) continue;
                    unsupportedTransform = algorithm;
                    return false;
                }
            }
            return true;
        }

        private static bool HasSupportedSignedInfoCanonicalizationMethod(
            XmlElement signatureElement,
            out string? unsupportedCanonicalization) {
            XmlElement? signedInfo = signatureElement.ChildNodes
                .OfType<XmlElement>()
                .FirstOrDefault(element =>
                    element.LocalName == "SignedInfo" &&
                    element.NamespaceURI == SignedXml.XmlDsigNamespaceUrl);
            XmlElement? canonicalization = signedInfo?.ChildNodes
                .OfType<XmlElement>()
                .FirstOrDefault(element =>
                    element.LocalName == "CanonicalizationMethod" &&
                    element.NamespaceURI == SignedXml.XmlDsigNamespaceUrl);
            string algorithm = canonicalization?.GetAttribute("Algorithm").Trim() ?? string.Empty;
            if (SupportedSignedInfoCanonicalizationMethods.Contains(algorithm)) {
                unsupportedCanonicalization = null;
                return true;
            }
            unsupportedCanonicalization = algorithm;
            return false;
        }

        private static AsymmetricAlgorithm? GetPublicKey(X509Certificate2 certificate) {
            AsymmetricAlgorithm? publicKey = certificate.GetRSAPublicKey();
            publicKey ??= certificate.GetECDsaPublicKey();
#if NETSTANDARD2_0 || NETFRAMEWORK
            publicKey ??= certificate.PublicKey.Key;
#else
            publicKey ??= certificate.GetDSAPublicKey();
#endif
            return publicKey;
        }

        private static bool HasOnlyLocalSignedInfoReferences(SignedXml signedXml, out string? unsupportedUri) {
            unsupportedUri = null;
            foreach (object item in signedXml.SignedInfo!.References) {
                if (item is not Reference reference) continue;
                string uri = reference.Uri ?? string.Empty;
                if (uri.Length > 0 && !uri.StartsWith("#", StringComparison.Ordinal)) {
                    unsupportedUri = uri;
                    return false;
                }
            }
            return true;
        }

        private static IEnumerable<X509Certificate2> ReadCertificates(
            XmlSignaturePart signaturePart,
            XmlDocument signatureXml,
            int maxCertificates,
            long maxCertificateBytes,
            OfficePackageCertificateByteBudget certificateByteBudget,
            List<WordSignatureValidationFinding> findings) {
            var result = new List<X509Certificate2>();
            IReadOnlyList<XmlElement> embedded = GetEmbeddedSignerCertificateElements(signatureXml);
            if (embedded.Count > maxCertificates) {
                throw new InvalidDataException("The XML signature exceeds the " + maxCertificates + " certificate limit.");
            }
            foreach (XmlElement element in embedded) {
                TryAddCertificate(element.InnerText, "embedded X509Certificate", maxCertificateBytes, certificateByteBudget, result, findings, signaturePart.Uri.ToString());
            }

            int declaredCertificateCount = embedded.Count;
            foreach (IdPartPair relationship in signaturePart.Parts) {
                OpenXmlPart relatedPart = relationship.OpenXmlPart;
                if (!IsCertificatePart(relatedPart)) continue;
                declaredCertificateCount++;
                if (declaredCertificateCount > maxCertificates) throw new InvalidDataException("The XML signature exceeds the " + maxCertificates + " certificate limit.");
                try {
                    using Stream stream = relatedPart.GetStream(FileMode.Open, FileAccess.Read);
                    if (stream.CanSeek && stream.Length > maxCertificateBytes) {
                        throw new InvalidDataException("The related signature certificate exceeds the " + maxCertificateBytes + " byte limit.");
                    }
                    using var buffer = new MemoryStream();
                    CopyBounded(stream, buffer, maxCertificateBytes);
                    byte[] certificateBytes = buffer.ToArray();
                    certificateByteBudget.Reserve(certificateBytes.LongLength);
                    result.Add(LoadCertificate(certificateBytes));
                } catch (InvalidDataException) {
                    throw;
                } catch (Exception exception) when (exception is IOException or CryptographicException or InvalidOperationException) {
                    findings.Add(Finding("CertificateMalformed", WordSignatureValidationState.Failed,
                        "The related signature certificate could not be read: " + exception.Message,
                        signaturePart.Uri.ToString()));
                }
            }
            return result;
        }

        private static IReadOnlyList<XmlElement> GetEmbeddedSignerCertificateElements(XmlDocument signatureXml) {
            XmlElement? signature = signatureXml.DocumentElement;
            if (signature == null) return Array.Empty<XmlElement>();
            return signature.ChildNodes
                .OfType<XmlElement>()
                .Where(element => element.LocalName == "KeyInfo" && element.NamespaceURI == SignedXml.XmlDsigNamespaceUrl)
                .SelectMany(keyInfo => keyInfo.ChildNodes.OfType<XmlElement>())
                .Where(element => element.LocalName == "X509Data" && element.NamespaceURI == SignedXml.XmlDsigNamespaceUrl)
                .SelectMany(x509Data => x509Data.ChildNodes.OfType<XmlElement>())
                .Where(element => element.LocalName == "X509Certificate" && element.NamespaceURI == SignedXml.XmlDsigNamespaceUrl)
                .ToArray();
        }

        private static void TryAddCertificate(
            string encoded,
            string source,
            long maxCertificateBytes,
            OfficePackageCertificateByteBudget certificateByteBudget,
            List<X509Certificate2> certificates,
            List<WordSignatureValidationFinding> findings,
            string signaturePartUri) {
            try {
                long maxEncodedCharacters = GetMaxBase64EncodedCharacters(maxCertificateBytes);
                if (encoded.Length > maxEncodedCharacters) {
                    throw new InvalidDataException("The " + source + " value exceeds the " + maxCertificateBytes + " byte limit.");
                }

                byte[] certificateBytes = Convert.FromBase64String(encoded);
                if (certificateBytes.LongLength > maxCertificateBytes) {
                    throw new InvalidDataException("The " + source + " value exceeds the " + maxCertificateBytes + " byte limit.");
                }
                certificateByteBudget.Reserve(certificateBytes.LongLength);
                certificates.Add(LoadCertificate(certificateBytes));
            } catch (Exception exception) when (exception is FormatException or CryptographicException) {
                findings.Add(Finding("CertificateMalformed", WordSignatureValidationState.Failed,
                    "The " + source + " value could not be decoded: " + exception.Message,
                    signaturePartUri));
            }
        }

        private static XmlElement? ReadSignatureValue(
            XmlDocument document,
            List<WordSignatureValidationFinding> findings,
            string signaturePartUri) {
            XmlElement? element = document.GetElementsByTagName("SignatureValue", SignedXml.XmlDsigNamespaceUrl)
                .OfType<XmlElement>()
                .FirstOrDefault();
            if (element == null) {
                findings.Add(Finding("SignatureValueMissing", WordSignatureValidationState.Failed,
                    "The XML signature does not contain SignatureValue.", signaturePartUri));
                return null;
            }
            try {
                Convert.FromBase64String(element.InnerText);
                return element;
            } catch (FormatException exception) {
                findings.Add(Finding("SignatureValueMalformed", WordSignatureValidationState.Failed,
                    "SignatureValue is not valid base64: " + exception.Message, signaturePartUri));
                return null;
            }
        }

        private static long GetMaxBase64EncodedCharacters(long maxDecodedBytes) {
            return maxDecodedBytes > (long.MaxValue / 4L) * 3L
                ? long.MaxValue
                : ((maxDecodedBytes + 2L) / 3L) * 4L;
        }

        private static XmlDocument LoadXml(byte[] bytes, long maxBytes) {
            if (bytes.LongLength > maxBytes) throw new InvalidDataException("The XML signature exceeds the " + maxBytes + " byte limit.");
            var settings = new XmlReaderSettings {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
                MaxCharactersInDocument = maxBytes
            };
            var document = new XmlDocument { PreserveWhitespace = true, XmlResolver = null };
            using var stream = new MemoryStream(bytes, writable: false);
            using XmlReader reader = XmlReader.Create(stream, settings);
            document.Load(reader);
            return document;
        }

        private static X509Certificate2 LoadCertificate(byte[] rawCertificate) {
#if NET9_0_OR_GREATER
            return X509CertificateLoader.LoadCertificate(rawCertificate);
#else
            return new X509Certificate2(rawCertificate);
#endif
        }

        private static bool IsCertificatePart(OpenXmlPart part) =>
            part.RelationshipType.EndsWith("/digital-signature/certificate", StringComparison.OrdinalIgnoreCase) ||
            part.Uri.ToString().EndsWith(".cer", StringComparison.OrdinalIgnoreCase);

        private static WordSignatureValidationState MapStatus(SecurityValidationStatus status) {
            switch (status) {
                case SecurityValidationStatus.Valid:
                    return WordSignatureValidationState.Passed;
                case SecurityValidationStatus.Invalid:
                    return WordSignatureValidationState.Failed;
                case SecurityValidationStatus.Indeterminate:
                    return WordSignatureValidationState.NotChecked;
                default:
                    return WordSignatureValidationState.NotChecked;
            }
        }

        private static WordSignatureValidationFinding Finding(
            string code,
            WordSignatureValidationState state,
            string message,
            string? signaturePartUri = null,
            string? referenceUri = null) =>
            new(code, state, message, signaturePartUri, referenceUri);

        private static WordSignaturePartValidationResult FailedMissingPart(WordSignaturePartInfo info) =>
            FailedMalformed(info, "The signature part described by package metadata could not be opened.");

        private static WordSignaturePartValidationResult FailedMalformed(WordSignaturePartInfo info, string message) =>
            new(
                info,
                WordSignatureValidationState.Failed,
                WordSignatureValidationState.NotChecked,
                WordSignatureValidationState.NotChecked,
                false,
                WordSignatureValidationState.NotChecked,
                null,
                Array.Empty<Rfc3161TimestampVerificationResult>(),
                new[] { Finding("SignaturePartMalformed", WordSignatureValidationState.Failed, message, info.Uri) });

        internal static void ValidateOptions(WordSignatureValidationOptions options) {
            if (options.MaxSignatureParts <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxSignatureParts));
            if (options.MaxPackageBytes <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxPackageBytes));
            if (options.MaxPackageParts <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxPackageParts));
            if (options.MaxPartBytes <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxPartBytes));
            if (options.MaxSignedReferences <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxSignedReferences));
            if (options.MaxTotalDigestBytes <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxTotalDigestBytes));
            if (options.MaxSignatureBytes <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxSignatureBytes));
            if (options.MaxCertificates <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxCertificates));
            if (options.MaxCertificateBytes <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxCertificateBytes));
            if (options.MaxTotalCertificateBytes <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxTotalCertificateBytes));
            if (options.MaxTimestampTokens <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxTimestampTokens));
            if (options.MaxTimestampBytes <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxTimestampBytes));
        }

        private static bool IsValidationException(Exception exception) =>
            exception is IOException or InvalidDataException or InvalidOperationException or ArgumentException or
                CryptographicException or XmlException or NotSupportedException;

        private static void CopyBounded(Stream source, Stream destination, long maxBytes) {
            byte[] buffer = new byte[81920];
            long total = 0;
            int read;
            while ((read = source.Read(buffer, 0, buffer.Length)) > 0) {
                total += read;
                if (total > maxBytes) throw new InvalidDataException("The related signature certificate exceeds the " + maxBytes + " byte limit.");
                destination.Write(buffer, 0, read);
            }
        }
    }
}
