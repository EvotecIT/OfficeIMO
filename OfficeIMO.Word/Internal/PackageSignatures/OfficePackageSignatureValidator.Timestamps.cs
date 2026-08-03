#nullable enable
using OfficeIMO.Security;
using System.Text;
using System.Xml;

namespace OfficeIMO.Word {
    internal static partial class OfficePackageSignatureValidator {
        private static void ValidateTimestampTokens(
            IReadOnlyList<XmlElement> tokens,
            XmlElement signatureValue,
            WordSignatureValidationOptions options,
            string signaturePartUri,
            IOfficeSecurityProvider securityProvider,
            OfficePackageTimestampValidationBudget timestampBudget,
            List<Rfc3161TimestampVerificationResult> results,
            List<WordSignatureValidationFinding> findings) {
            if (tokens.Count > options.MaxTimestampTokens) {
                throw new InvalidDataException("The XML signature exceeds the " + options.MaxTimestampTokens + " timestamp-token limit.");
            }
            foreach (XmlElement tokenElement in tokens) {
                timestampBudget.ReserveToken();
                byte[] encoded;
                try {
                    if (OfficePackageBase64.ExceedsDecodedByteLimit(
                        tokenElement.InnerText,
                        options.MaxTimestampBytes)) {
                        throw new InvalidDataException("An embedded RFC 3161 timestamp token exceeds the " + options.MaxTimestampBytes + " byte limit.");
                    }
                    encoded = Convert.FromBase64String(tokenElement.InnerText);
                    if (encoded.LongLength > options.MaxTimestampBytes) {
                        throw new InvalidDataException("An embedded RFC 3161 timestamp token exceeds the " + options.MaxTimestampBytes + " byte limit.");
                    }
                    timestampBudget.ReserveBytes(encoded.LongLength);
                } catch (FormatException exception) {
                    findings.Add(Finding("TimestampMalformed", WordSignatureValidationState.Failed,
                        "An embedded RFC 3161 timestamp token is not valid base64: " + exception.Message,
                        signaturePartUri));
                    continue;
                }
                byte[] timestampedData;
                try {
                    timestampedData = CanonicalizeTimestampedSignatureValue(
                        signatureValue,
                        tokenElement,
                        securityProvider,
                        options.MaxSignatureBytes);
                } catch (NotSupportedException exception) {
                    findings.Add(Finding("TimestampCanonicalizationUnsupported", WordSignatureValidationState.Unsupported,
                        exception.Message, signaturePartUri));
                    continue;
                }
                timestampBudget.ReserveVerification();
                Rfc3161TimestampVerificationResult result = securityProvider.VerifyTimestamp(
                    encoded,
                    timestampedData,
                    options.TimestampCertificateValidation,
                    options.MaxTimestampBytes,
                    options.MaxCertificates);
                results.Add(result);
                foreach (SecurityFinding finding in result.Findings) {
                    findings.Add(Finding(finding.Code, MapStatus(result.Status), finding.Message, signaturePartUri));
                }
            }
        }

        private sealed class OfficePackageTimestampValidationBudget {
            private readonly int _maxTokens;
            private readonly long _maxBytes;
            private int _tokens;
            private int _verifications;
            private long _bytes;

            internal OfficePackageTimestampValidationBudget(int maxTokens, long maxTimestampBytes) {
                _maxTokens = maxTokens;
                _maxBytes = maxTimestampBytes > long.MaxValue / maxTokens
                    ? long.MaxValue
                    : maxTimestampBytes * maxTokens;
            }

            internal void ReserveToken() {
                if (_tokens >= _maxTokens) {
                    throw new InvalidDataException("The validation operation exceeds the " + _maxTokens + " aggregate timestamp-token limit.");
                }
                _tokens++;
            }

            internal void ReserveBytes(long byteCount) {
                if (byteCount < 0 || _bytes > _maxBytes - byteCount) {
                    throw new InvalidDataException("The validation operation exceeds the " + _maxBytes + " byte aggregate timestamp-token limit.");
                }
                _bytes += byteCount;
            }

            internal void ReserveVerification() {
                if (_verifications >= _maxTokens) {
                    throw new InvalidDataException("The validation operation exceeds the " + _maxTokens + " aggregate timestamp-verification limit.");
                }
                _verifications++;
            }
        }

        private static IReadOnlyList<XmlElement> GetXadesTimestampTokens(XmlElement signatureElement) {
            string signatureId = signatureElement.GetAttribute("Id");
            if (string.IsNullOrWhiteSpace(signatureId)) return Array.Empty<XmlElement>();
            string target = "#" + signatureId;

            return signatureElement.ChildNodes
                .OfType<XmlElement>()
                .Where(element => element.LocalName == "Object" && element.NamespaceURI == XmlDigitalSignatureAlgorithms.Namespace)
                .SelectMany(dataObject => dataObject.ChildNodes.OfType<XmlElement>())
                .Where(qualifyingProperties =>
                    qualifyingProperties.LocalName == "QualifyingProperties" &&
                    XadesNamespaces.Contains(qualifyingProperties.NamespaceURI) &&
                    string.Equals(qualifyingProperties.GetAttribute("Target"), target, StringComparison.Ordinal))
                .SelectMany(qualifyingProperties => qualifyingProperties.ChildNodes
                    .OfType<XmlElement>()
                    .Where(element =>
                        element.LocalName == "UnsignedProperties" &&
                        element.NamespaceURI == qualifyingProperties.NamespaceURI))
                .SelectMany(unsignedProperties => unsignedProperties.ChildNodes
                    .OfType<XmlElement>()
                    .Where(element =>
                        element.LocalName == "UnsignedSignatureProperties" &&
                        element.NamespaceURI == unsignedProperties.NamespaceURI))
                .SelectMany(unsignedSignatureProperties => unsignedSignatureProperties.ChildNodes
                    .OfType<XmlElement>()
                    .Where(element =>
                        element.LocalName == "SignatureTimeStamp" &&
                        element.NamespaceURI == unsignedSignatureProperties.NamespaceURI))
                .SelectMany(signatureTimeStamp => signatureTimeStamp.ChildNodes
                    .OfType<XmlElement>()
                    .Where(element =>
                        element.LocalName == "EncapsulatedTimeStamp" &&
                        element.NamespaceURI == signatureTimeStamp.NamespaceURI))
                .ToArray();
        }

        private static byte[] CanonicalizeTimestampedSignatureValue(
            XmlElement signatureValue,
            XmlElement tokenElement,
            IOfficeSecurityProvider securityProvider,
            long maxOutputBytes) {
            XmlElement? timestampProperty = tokenElement;
            while (timestampProperty != null && !timestampProperty.LocalName.Equals("SignatureTimeStamp", StringComparison.Ordinal)) {
                timestampProperty = timestampProperty.ParentNode as XmlElement;
            }
            XmlElement? canonicalizationMethod = timestampProperty?
                .ChildNodes
                .OfType<XmlElement>()
                .FirstOrDefault(element =>
                    element.LocalName.Equals("CanonicalizationMethod", StringComparison.Ordinal) &&
                    element.NamespaceURI == XmlDigitalSignatureAlgorithms.Namespace);
            string algorithm = canonicalizationMethod?.GetAttribute("Algorithm") ?? XmlDigitalSignatureAlgorithms.CanonicalXml;
            string? inclusiveNamespacesPrefixList = canonicalizationMethod?
                .ChildNodes
                .OfType<XmlElement>()
                .FirstOrDefault(element =>
                    element.LocalName.Equals("InclusiveNamespaces", StringComparison.Ordinal) &&
                    element.NamespaceURI == XmlDigitalSignatureAlgorithms.ExclusiveCanonicalXml)?
                .GetAttribute("PrefixList");

            var input = new XmlDocument { PreserveWhitespace = true, XmlResolver = null };
            XmlElement imported = (XmlElement)input.ImportNode(signatureValue, deep: true);
            var namespaceNames = new HashSet<string>(StringComparer.Ordinal);
            var inheritedXmlAttributeNames = new HashSet<string>(StringComparer.Ordinal);
            bool includeInheritedXmlAttributes =
                algorithm == XmlDigitalSignatureAlgorithms.CanonicalXml ||
                algorithm == XmlDigitalSignatureAlgorithms.CanonicalXmlWithComments;
            for (XmlElement? ancestor = signatureValue; ancestor != null; ancestor = ancestor.ParentNode as XmlElement) {
                foreach (XmlAttribute attribute in ancestor.Attributes) {
                    if (attribute.Prefix == "xmlns" || attribute.Name == "xmlns") {
                        if (!namespaceNames.Add(attribute.Name) || imported.HasAttribute(attribute.Name)) continue;
                        imported.Attributes.Append((XmlAttribute)input.ImportNode(attribute, deep: true));
                        continue;
                    }
                    if (!includeInheritedXmlAttributes ||
                        attribute.NamespaceURI != "http://www.w3.org/XML/1998/namespace") {
                        continue;
                    }
                    string attributeKey = attribute.NamespaceURI + "\0" + attribute.LocalName;
                    if (!inheritedXmlAttributeNames.Add(attributeKey) ||
                        imported.HasAttribute(attribute.LocalName, attribute.NamespaceURI)) {
                        continue;
                    }
                    imported.Attributes.Append((XmlAttribute)input.ImportNode(attribute, deep: true));
                }
            }
            input.AppendChild(imported);
            return securityProvider.CanonicalizeXml(
                Encoding.UTF8.GetBytes(input.OuterXml),
                algorithm,
                inclusiveNamespacesPrefixList,
                maxOutputBytes);
        }

        private static WordSignatureValidationState ResolveTimestampStatus(
            XmlDocument document,
            int declaredTokenCount,
            WordSignatureValidationOptions options,
            IReadOnlyList<Rfc3161TimestampVerificationResult> timestampResults,
            string signaturePartUri,
            List<WordSignatureValidationFinding> findings) {
            if (!options.ValidateTimestamps) return WordSignatureValidationState.NotChecked;
            if (timestampResults.Any(result => result.Status == SecurityValidationStatus.Invalid)) {
                return WordSignatureValidationState.Failed;
            }
            if (declaredTokenCount > timestampResults.Count) {
                if (findings.Any(finding => finding.Code == "TimestampMalformed")) {
                    findings.Add(Finding("TimestampValidationFailed", WordSignatureValidationState.Failed,
                        "At least one embedded RFC 3161 timestamp token could not be decoded or validated.",
                        signaturePartUri));
                    return WordSignatureValidationState.Failed;
                }
                if (findings.Any(finding => finding.Code == "TimestampCanonicalizationUnsupported")) {
                    return WordSignatureValidationState.Unsupported;
                }
                findings.Add(Finding("TimestampValidationFailed", WordSignatureValidationState.Failed,
                    "At least one embedded RFC 3161 timestamp token could not be decoded or validated.",
                    signaturePartUri));
                return WordSignatureValidationState.Failed;
            }
            if (timestampResults.Count > 0) {
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
                DisableCertificateDownloads = source.DisableCertificateDownloads,
                VerificationTime = verificationTime,
                UrlRetrievalTimeout = source.UrlRetrievalTimeout,
                ChainEvaluator = source.ChainEvaluator
            };
            result.ExtraCertificates.AddRange(source.ExtraCertificates);
            return result;
        }
    }
}
