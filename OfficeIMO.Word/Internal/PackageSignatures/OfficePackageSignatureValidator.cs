#nullable enable
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Security;
using System.Collections;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Security.Cryptography.Xml;
using System.Xml;

namespace OfficeIMO.Word {
    /// <summary>Cross-platform XML DSig, X.509, revocation, and RFC 3161 validator for OPC signatures.</summary>
    internal static class OfficePackageSignatureValidator {
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
                results.Add(ValidateSignaturePart(signaturePart, signaturePartInfo, archive, options));
            }
            return results;
        }

        private static WordSignaturePartValidationResult ValidateSignaturePart(
            XmlSignaturePart signaturePart,
            WordSignaturePartInfo signaturePartInfo,
            OfficePackageSignatureArchive archive,
            WordSignatureValidationOptions options) {
            var findings = new List<WordSignatureValidationFinding>();
            var timestampResults = new List<Rfc3161TimestampVerificationResult>();
            var certificates = new List<X509Certificate2>();
            try {
                byte[] signatureBytes = archive.ReadPart(signaturePartInfo.Uri, options.MaxSignatureBytes);
                XmlDocument document = LoadXml(signatureBytes, options.MaxSignatureBytes);
                XmlElement? signatureElement = document.DocumentElement;
                if (signatureElement == null ||
                    signatureElement.LocalName != "Signature" ||
                    signatureElement.NamespaceURI != SignedXml.XmlDsigNamespaceUrl) {
                    return FailedMalformed(signaturePartInfo, "The signature part does not contain an XML DSig Signature root element.");
                }

                certificates.AddRange(ReadCertificates(
                    signaturePart,
                    document,
                    options.MaxCertificates,
                    options.MaxCertificateBytes,
                    findings));
                byte[]? signatureValue = ReadSignatureValue(document, findings, signaturePartInfo.Uri);
                WordSignatureValidationState cryptographicStatus;
                X509Certificate2? signer = null;

                if (!options.ValidateCryptographicSignature) {
                    cryptographicStatus = WordSignatureValidationState.NotChecked;
                    findings.Add(Finding("CryptographicValidationDisabled", cryptographicStatus,
                        "XML DSig signature-value validation was disabled by caller policy.", signaturePartInfo.Uri));
                } else if (certificates.Count == 0) {
                    cryptographicStatus = WordSignatureValidationState.Unsupported;
                    findings.Add(Finding("SignerCertificateMissing", cryptographicStatus,
                        "No embedded or related X.509 signer certificate was found.", signaturePartInfo.Uri));
                } else {
                    cryptographicStatus = ValidateSignedXml(document, signatureElement, certificates, signaturePartInfo.Uri, findings, out signer);
                }

                if (options.ValidateTimestamps && signatureValue != null) {
                    ValidateTimestampTokens(
                        document,
                        signatureValue,
                        options,
                        signaturePartInfo.Uri,
                        timestampResults,
                        findings);
                }

                WordSignatureValidationState timestampStatus = ResolveTimestampStatus(
                    document,
                    options,
                    timestampResults,
                    signaturePartInfo.Uri,
                    findings);

                CertificateValidationResult? certificateValidation = null;
                WordSignatureValidationState certificateStatus;
                WordSignatureValidationState revocationStatus;
                if (signer == null) {
                    certificateStatus = certificates.Count == 0
                        ? WordSignatureValidationState.NotPresent
                        : WordSignatureValidationState.NotChecked;
                    revocationStatus = certificateStatus;
                } else {
                    CertificateValidationOptions signerOptions = ResolveSignerCertificateValidation(
                        options.CertificateValidation,
                        timestampResults);
                    CertificateTrustValidationResult trust = CertificateValidator.Validate(
                        signer,
                        certificates.Where(certificate => !ReferenceEquals(certificate, signer)),
                        signerOptions,
                        CertificateValidationPurpose.DocumentSigning);
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
                    timestampStatus,
                    certificateValidation,
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
                    WordSignatureValidationState.NotChecked,
                    null,
                    timestampResults.ToArray(),
                    findings.ToArray());
            } finally {
                foreach (X509Certificate2 certificate in certificates) certificate.Dispose();
            }
        }

        private static WordSignatureValidationState ValidateSignedXml(
            XmlDocument document,
            XmlElement signatureElement,
            IReadOnlyList<X509Certificate2> certificates,
            string signaturePartUri,
            List<WordSignatureValidationFinding> findings,
            out X509Certificate2? signer) {
            signer = null;
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

            foreach (X509Certificate2 candidate in certificates) {
                try {
                    if (!signedXml.CheckSignature(candidate, verifySignatureOnly: true)) continue;
                    signer = candidate;
                    findings.Add(Finding("XmlSignatureValid", WordSignatureValidationState.Passed,
                        "XML DSig signature-value and signed-object validation passed.", signaturePartUri));
                    return WordSignatureValidationState.Passed;
                } catch (CryptographicException) {
                    // Try the remaining embedded or related certificate candidates.
                }
            }

            findings.Add(Finding("XmlSignatureInvalid", WordSignatureValidationState.Failed,
                "XML DSig signature-value or signed-object validation failed for every supplied certificate.", signaturePartUri));
            return WordSignatureValidationState.Failed;
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
            List<WordSignatureValidationFinding> findings) {
            var result = new List<X509Certificate2>();
            XmlNodeList embedded = signatureXml.GetElementsByTagName("X509Certificate", SignedXml.XmlDsigNamespaceUrl);
            foreach (XmlElement element in embedded.OfType<XmlElement>()) {
                if (result.Count >= maxCertificates) throw new InvalidDataException("The XML signature exceeds the " + maxCertificates + " certificate limit.");
                TryAddCertificate(element.InnerText, "embedded X509Certificate", result, findings, signaturePart.Uri.ToString());
            }

            foreach (IdPartPair relationship in signaturePart.Parts) {
                OpenXmlPart relatedPart = relationship.OpenXmlPart;
                if (!IsCertificatePart(relatedPart)) continue;
                if (result.Count >= maxCertificates) throw new InvalidDataException("The XML signature exceeds the " + maxCertificates + " certificate limit.");
                try {
                    using Stream stream = relatedPart.GetStream(FileMode.Open, FileAccess.Read);
                    if (stream.CanSeek && stream.Length > maxCertificateBytes) {
                        throw new InvalidDataException("The related signature certificate exceeds the " + maxCertificateBytes + " byte limit.");
                    }
                    using var buffer = new MemoryStream();
                    CopyBounded(stream, buffer, maxCertificateBytes);
                    result.Add(LoadCertificate(buffer.ToArray()));
                } catch (Exception exception) when (exception is IOException or CryptographicException or InvalidOperationException) {
                    findings.Add(Finding("CertificateMalformed", WordSignatureValidationState.Failed,
                        "The related signature certificate could not be read: " + exception.Message,
                        signaturePart.Uri.ToString()));
                }
            }
            return result;
        }

        private static void TryAddCertificate(
            string encoded,
            string source,
            List<X509Certificate2> certificates,
            List<WordSignatureValidationFinding> findings,
            string signaturePartUri) {
            try {
                certificates.Add(LoadCertificate(Convert.FromBase64String(encoded)));
            } catch (Exception exception) when (exception is FormatException or CryptographicException) {
                findings.Add(Finding("CertificateMalformed", WordSignatureValidationState.Failed,
                    "The " + source + " value could not be decoded: " + exception.Message,
                    signaturePartUri));
            }
        }

        private static void ValidateTimestampTokens(
            XmlDocument document,
            byte[] signatureValue,
            WordSignatureValidationOptions options,
            string signaturePartUri,
            List<Rfc3161TimestampVerificationResult> results,
            List<WordSignatureValidationFinding> findings) {
            XmlNodeList tokens = document.GetElementsByTagName("EncapsulatedTimeStamp", "*");
            if (tokens.Count > options.MaxTimestampTokens) {
                throw new InvalidDataException("The XML signature exceeds the " + options.MaxTimestampTokens + " timestamp-token limit.");
            }
            foreach (XmlElement tokenElement in tokens.OfType<XmlElement>()) {
                byte[] encoded;
                try {
                    encoded = Convert.FromBase64String(tokenElement.InnerText);
                } catch (FormatException exception) {
                    findings.Add(Finding("TimestampMalformed", WordSignatureValidationState.Failed,
                        "An embedded RFC 3161 timestamp token is not valid base64: " + exception.Message,
                        signaturePartUri));
                    continue;
                }
                Rfc3161TimestampVerificationResult result = Rfc3161TimestampVerifier.Verify(
                    encoded,
                    signatureValue,
                    options.TimestampCertificateValidation,
                    options.MaxTimestampBytes,
                    options.MaxCertificates);
                results.Add(result);
                foreach (SecurityFinding finding in result.Findings) {
                    findings.Add(Finding(finding.Code, MapStatus(result.Status), finding.Message, signaturePartUri));
                }
            }
        }

        private static WordSignatureValidationState ResolveTimestampStatus(
            XmlDocument document,
            WordSignatureValidationOptions options,
            IReadOnlyList<Rfc3161TimestampVerificationResult> timestampResults,
            string signaturePartUri,
            List<WordSignatureValidationFinding> findings) {
            if (!options.ValidateTimestamps) return WordSignatureValidationState.NotChecked;
            int declaredTokenCount = document.GetElementsByTagName("EncapsulatedTimeStamp", "*").Count;
            if (declaredTokenCount > timestampResults.Count) {
                findings.Add(Finding("TimestampValidationFailed", WordSignatureValidationState.Failed,
                    "At least one embedded RFC 3161 timestamp token could not be decoded or validated.",
                    signaturePartUri));
                return WordSignatureValidationState.Failed;
            }
            if (timestampResults.Count > 0) {
                if (timestampResults.Any(result => result.Status == SecurityValidationStatus.Invalid)) return WordSignatureValidationState.Failed;
                if (timestampResults.All(result => result.Status == SecurityValidationStatus.Valid)) return WordSignatureValidationState.Passed;
                return WordSignatureValidationState.Unsupported;
            }

            bool hasClaimedTime = document.GetElementsByTagName("SignatureTime", "*").Count > 0 ||
                                  document.GetElementsByTagName("SigningTime", "*").Count > 0;
            if (hasClaimedTime) {
                findings.Add(Finding("ClaimedSigningTimeNotTrusted", WordSignatureValidationState.NotPresent,
                    "The signature contains a claimed signing time but no RFC 3161 timestamp-authority token.",
                    signaturePartUri));
            }
            return WordSignatureValidationState.NotPresent;
        }

        private static CertificateValidationOptions ResolveSignerCertificateValidation(
            CertificateValidationOptions source,
            IReadOnlyList<Rfc3161TimestampVerificationResult> timestamps) {
            DateTime? verificationTime = source.VerificationTime;
            if (verificationTime == null) {
                verificationTime = timestamps
                    .Where(result => result.Status == SecurityValidationStatus.Valid && result.Timestamp.HasValue)
                    .Select(result => (DateTime?)result.Timestamp!.Value.UtcDateTime)
                    .OrderBy(value => value)
                    .FirstOrDefault();
            }

            var result = new CertificateValidationOptions {
                ValidateChain = source.ValidateChain,
                RevocationMode = source.RevocationMode,
                RevocationFlag = source.RevocationFlag,
                VerificationFlags = source.VerificationFlags,
                VerificationTime = verificationTime,
                UrlRetrievalTimeout = source.UrlRetrievalTimeout,
                ChainEvaluator = source.ChainEvaluator
            };
            result.ExtraCertificates.AddRange(source.ExtraCertificates);
            return result;
        }

        private static byte[]? ReadSignatureValue(
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
                return Convert.FromBase64String(element.InnerText);
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
                WordSignatureValidationState.NotChecked,
                null,
                Array.Empty<Rfc3161TimestampVerificationResult>(),
                new[] { Finding("SignaturePartMalformed", WordSignatureValidationState.Failed, message, info.Uri) });

        internal static void ValidateOptions(WordSignatureValidationOptions options) {
            if (options.MaxSignatureParts <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxSignatureParts));
            if (options.MaxPackageBytes <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxPackageBytes));
            if (options.MaxPackageParts <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxPackageParts));
            if (options.MaxPartBytes <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxPartBytes));
            if (options.MaxSignatureBytes <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxSignatureBytes));
            if (options.MaxCertificates <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxCertificates));
            if (options.MaxCertificateBytes <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxCertificateBytes));
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
