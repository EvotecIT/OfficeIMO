#nullable enable
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Security;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
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
            XmlDigitalSignatureAlgorithms.CanonicalXml,
            XmlDigitalSignatureAlgorithms.CanonicalXmlWithComments,
            XmlDigitalSignatureAlgorithms.ExclusiveCanonicalXml,
            XmlDigitalSignatureAlgorithms.ExclusiveCanonicalXmlWithComments,
            XmlDigitalSignatureAlgorithms.EnvelopedSignatureTransform
        };
        private static readonly HashSet<string> SupportedSignedInfoCanonicalizationMethods = new(StringComparer.Ordinal) {
            XmlDigitalSignatureAlgorithms.CanonicalXml,
            XmlDigitalSignatureAlgorithms.CanonicalXmlWithComments,
            XmlDigitalSignatureAlgorithms.ExclusiveCanonicalXml,
            XmlDigitalSignatureAlgorithms.ExclusiveCanonicalXmlWithComments
        };

        internal static IReadOnlyList<WordSignaturePartValidationResult> Validate(
            DigitalSignatureOriginPart? originPart,
            byte[] packageBytes,
            WordSignatureInfo signatureInfo,
            IOfficeSecurityProvider securityProvider,
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
            using var archive = new OfficePackageSignatureArchive(
                packageBytes,
                options.MaxPackageParts,
                options.MaxPartBytes,
                securityProvider);
            var certificateByteBudget = new OfficePackageCertificateByteBudget(options.MaxTotalCertificateBytes);
            var timestampBudget = new OfficePackageTimestampValidationBudget(
                options.MaxTimestampTokens,
                options.MaxTimestampBytes);
            var digestWorkBudget = new OfficePackageDigestWorkBudget(options.MaxTotalDigestBytes);
            digestWorkBudget.Reserve(signatureInfo.InspectionDigestWorkBytes);
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
                    securityProvider,
                    options,
                    certificateByteBudget,
                    timestampBudget,
                    digestWorkBudget));
            }
            return results;
        }

        private static WordSignaturePartValidationResult ValidateSignaturePart(
            XmlSignaturePart signaturePart,
            WordSignaturePartInfo signaturePartInfo,
            OfficePackageSignatureArchive archive,
            IOfficeSecurityProvider securityProvider,
            WordSignatureValidationOptions options,
            OfficePackageCertificateByteBudget certificateByteBudget,
            OfficePackageTimestampValidationBudget timestampBudget,
            OfficePackageDigestWorkBudget digestWorkBudget) {
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
                    signatureElement.NamespaceURI != XmlDigitalSignatureAlgorithms.Namespace) {
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
                AddCallerCertificateCandidates(
                    options.CertificateValidation.ExtraCertificates,
                    options.MaxCertificates,
                    options.MaxCertificateBytes,
                    certificateByteBudget,
                    certificates);
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
                        "No embedded, related, or caller-supplied X.509 signer certificate was found.", signaturePartInfo.Uri));
                } else if (!ValidateSignedInfoPolicy(
                               signatureElement,
                               signaturePartInfo.Uri,
                               findings)) {
                    cryptographicStatus = WordSignatureValidationState.Unsupported;
                } else {
                    long availableDigestWorkBytes = digestWorkBudget.RemainingBytes;
                    if (availableDigestWorkBytes <= 0) {
                        throw new InvalidDataException(
                            "Local SignedInfo references exceed the " + digestWorkBudget.MaxBytes +
                            " byte aggregate digest-work limit across signature parts.");
                    }
                    long localDigestWorkBytes = XmlDigitalSignatureReferenceWorkCalculator.Measure(
                        document,
                        signatureElement.ChildNodes
                            .OfType<XmlElement>()
                            .FirstOrDefault(element =>
                                element.LocalName == "SignedInfo" &&
                                element.NamespaceURI == XmlDigitalSignatureAlgorithms.Namespace),
                        certificates.Count,
                        availableDigestWorkBytes);
                    digestWorkBudget.Reserve(localDigestWorkBytes);
                    var verificationRequest = new XmlDigitalSignatureVerificationRequest(
                        signatureBytes,
                        certificates) {
                        MaxSignatureBytes = options.MaxSignatureBytes,
                        MaxReferences = options.MaxSignedReferences,
                        MaxTotalDigestWorkBytes = availableDigestWorkBytes,
                        AllowedCanonicalizationMethods = SupportedSignedInfoCanonicalizationMethods,
                        AllowedReferenceTransforms = SupportedSignedInfoReferenceTransforms
                    };
                    XmlDigitalSignatureVerificationResult verification =
                        securityProvider.VerifyXmlSignature(verificationRequest);
                    cryptographicStatus = MapStatus(verification.Status);
                    matchingSigners = verification.MatchingCertificates;
                    foreach (SecurityFinding finding in verification.Findings) {
                        findings.Add(Finding(
                            finding.Code,
                            MapStatus(verification.Status),
                            finding.Message,
                            signaturePartInfo.Uri));
                    }
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
                            securityProvider,
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
                        securityProvider,
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

        private static CertificateTrustValidationResult SelectSignerTrust(
            IReadOnlyList<X509Certificate2> matchingSigners,
            IReadOnlyList<X509Certificate2> certificates,
            CertificateValidationOptions options,
            IOfficeSecurityProvider securityProvider,
            bool revocationCheckRequired) {
            CertificateTrustValidationResult? fallback = null;
            foreach (X509Certificate2 signer in matchingSigners) {
                CertificateTrustValidationResult trust = securityProvider.ValidateCertificate(
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

        private static IEnumerable<X509Certificate2> ReadCertificates(
            XmlSignaturePart signaturePart,
            XmlDocument signatureXml,
            int maxCertificates,
            long maxCertificateBytes,
            OfficePackageCertificateByteBudget certificateByteBudget,
            List<WordSignatureValidationFinding> findings) {
            var result = new List<X509Certificate2>();
            try {
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
            } catch {
                foreach (X509Certificate2 certificate in result) certificate.Dispose();
                throw;
            }
        }

        private static void AddCallerCertificateCandidates(
            X509Certificate2Collection callerCertificates,
            int maxCertificates,
            long maxCertificateBytes,
            OfficePackageCertificateByteBudget certificateByteBudget,
            List<X509Certificate2> certificates) {
            var identities = new HashSet<string>(
                certificates.Select(GetCertificateCandidateIdentity),
                StringComparer.OrdinalIgnoreCase);
            foreach (X509Certificate2 candidate in callerCertificates) {
                string identity = GetCertificateCandidateIdentity(candidate);
                if (!identities.Add(identity)) continue;
                if (certificates.Count >= maxCertificates) {
                    throw new InvalidDataException(
                        "The XML signature exceeds the " + maxCertificates + " certificate limit.");
                }
                byte[] rawCertificate = candidate.RawData;
                if (rawCertificate.LongLength > maxCertificateBytes) {
                    throw new InvalidDataException(
                        "A caller-supplied signature certificate exceeds the " +
                        maxCertificateBytes + " byte limit.");
                }
                certificateByteBudget.Reserve(rawCertificate.LongLength);
                certificates.Add(LoadCertificate(rawCertificate));
            }
        }

        private static string GetCertificateCandidateIdentity(X509Certificate2 certificate) =>
            certificate.Thumbprint ?? Convert.ToBase64String(certificate.RawData);

        private static IReadOnlyList<XmlElement> GetEmbeddedSignerCertificateElements(XmlDocument signatureXml) {
            XmlElement? signature = signatureXml.DocumentElement;
            if (signature == null) return Array.Empty<XmlElement>();
            return signature.ChildNodes
                .OfType<XmlElement>()
                .Where(element => element.LocalName == "KeyInfo" && element.NamespaceURI == XmlDigitalSignatureAlgorithms.Namespace)
                .SelectMany(keyInfo => keyInfo.ChildNodes.OfType<XmlElement>())
                .Where(element => element.LocalName == "X509Data" && element.NamespaceURI == XmlDigitalSignatureAlgorithms.Namespace)
                .SelectMany(x509Data => x509Data.ChildNodes.OfType<XmlElement>())
                .Where(element => element.LocalName == "X509Certificate" && element.NamespaceURI == XmlDigitalSignatureAlgorithms.Namespace)
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
                if (OfficePackageBase64.ExceedsDecodedByteLimit(encoded, maxCertificateBytes)) {
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
            XmlElement? element = document.GetElementsByTagName("SignatureValue", XmlDigitalSignatureAlgorithms.Namespace)
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

        private static bool ValidateSignedInfoPolicy(
            XmlElement signatureElement,
            string signaturePartUri,
            ICollection<WordSignatureValidationFinding> findings) {
            XmlElement? signedInfo = signatureElement.ChildNodes
                .OfType<XmlElement>()
                .FirstOrDefault(element =>
                    element.LocalName == "SignedInfo" &&
                    element.NamespaceURI == XmlDigitalSignatureAlgorithms.Namespace);
            if (signedInfo == null) return true;

            XmlElement[] references = GetSignedInfoReferences(signedInfo);

            XmlElement? canonicalization = signedInfo.ChildNodes
                .OfType<XmlElement>()
                .FirstOrDefault(element =>
                    element.LocalName == "CanonicalizationMethod" &&
                    element.NamespaceURI == XmlDigitalSignatureAlgorithms.Namespace);
            string canonicalizationAlgorithm = canonicalization?.GetAttribute("Algorithm") ?? string.Empty;
            if (canonicalization != null &&
                !SupportedSignedInfoCanonicalizationMethods.Contains(canonicalizationAlgorithm)) {
                findings.Add(Finding(
                    "UnsupportedSignedInfoCanonicalizationMethod",
                    WordSignatureValidationState.Unsupported,
                    "SignedInfo canonicalization method '" + canonicalizationAlgorithm + "' is outside caller policy.",
                    signaturePartUri));
                return false;
            }

            foreach (XmlElement reference in references) {
                string uri = reference.GetAttribute("URI");
                if (uri.Length > 0 && uri[0] != '#') {
                    findings.Add(Finding(
                        "ExternalSignedInfoReference",
                        WordSignatureValidationState.Unsupported,
                        "SignedInfo reference '" + uri + "' is not a local fragment and was not dereferenced.",
                        signaturePartUri));
                    return false;
                }
                foreach (XmlElement transform in reference.ChildNodes
                             .OfType<XmlElement>()
                             .Where(element =>
                                 element.LocalName == "Transforms" &&
                                 element.NamespaceURI == XmlDigitalSignatureAlgorithms.Namespace)
                             .SelectMany(element => element.ChildNodes.OfType<XmlElement>())
                             .Where(element =>
                                 element.LocalName == "Transform" &&
                                 element.NamespaceURI == XmlDigitalSignatureAlgorithms.Namespace)) {
                    string algorithm = transform.GetAttribute("Algorithm");
                    if (SupportedSignedInfoReferenceTransforms.Contains(algorithm)) continue;
                    findings.Add(Finding(
                        "UnsupportedSignedInfoTransform",
                        WordSignatureValidationState.Unsupported,
                        "SignedInfo reference transform '" + algorithm + "' is outside caller policy.",
                        signaturePartUri));
                    return false;
                }
            }
            return true;
        }

        private static void EnsureSignedInfoReferenceCountWithinLimit(
            XmlElement signatureElement,
            int maxSignedReferences) {
            XmlElement? signedInfo = signatureElement.ChildNodes
                .OfType<XmlElement>()
                .FirstOrDefault(element =>
                    element.LocalName == "SignedInfo" &&
                    element.NamespaceURI == XmlDigitalSignatureAlgorithms.Namespace);
            int referenceCount = signedInfo == null ? 0 : GetSignedInfoReferences(signedInfo).Length;
            if (referenceCount > maxSignedReferences) {
                throw new InvalidDataException(
                    "The XML signature contains more than " + maxSignedReferences + " SignedInfo references.");
            }
        }

        private static XmlElement[] GetSignedInfoReferences(XmlElement signedInfo) =>
            signedInfo.ChildNodes
                .OfType<XmlElement>()
                .Where(element =>
                    element.LocalName == "Reference" &&
                    element.NamespaceURI == XmlDigitalSignatureAlgorithms.Namespace)
                .ToArray();


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
