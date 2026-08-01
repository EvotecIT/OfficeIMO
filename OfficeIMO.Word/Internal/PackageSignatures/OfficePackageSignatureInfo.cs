#nullable enable
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Xml.Linq;

namespace OfficeIMO.Word {
    internal enum OfficePackageSignatureDigestVerificationStatus {
        NotChecked,
        Passed,
        Failed,
        Unsupported
    }

    /// <summary>
    /// Describes package-level digital-signature metadata found in an Open XML package.
    /// </summary>
    internal sealed class OfficePackageSignatureInfo {
        internal OfficePackageSignatureInfo(
            bool hasDigitalSignatureOriginPart,
            string? originPartUri,
            string? originRelationshipId,
            bool hasApplicationSignatureMetadata,
            IReadOnlyList<OfficePackageSignaturePartInfo> signatureParts,
            IReadOnlyList<string> unsupportedDetails,
            IReadOnlyList<string> details) {
            HasDigitalSignatureOriginPart = hasDigitalSignatureOriginPart;
            OriginPartUri = originPartUri;
            OriginRelationshipId = originRelationshipId;
            HasApplicationSignatureMetadata = hasApplicationSignatureMetadata;
            SignatureParts = signatureParts;
            UnsupportedDetails = unsupportedDetails;
            Details = details;
        }

        /// <summary>Gets whether any package signature metadata was found.</summary>
        public bool HasSignatures => HasDigitalSignatureOriginPart || HasApplicationSignatureMetadata || SignatureParts.Count > 0;

        /// <summary>Gets whether the package contains a digital-signature origin part.</summary>
        public bool HasDigitalSignatureOriginPart { get; }

        /// <summary>Gets the signature origin part URI when present.</summary>
        public string? OriginPartUri { get; }

        /// <summary>Gets the package relationship id for the signature origin part when available.</summary>
        public string? OriginRelationshipId { get; }

        /// <summary>Gets whether application properties contain digital-signature metadata.</summary>
        public bool HasApplicationSignatureMetadata { get; }

        /// <summary>Gets XML signature parts discovered under the signature origin.</summary>
        public IReadOnlyList<OfficePackageSignaturePartInfo> SignatureParts { get; }

        /// <summary>Gets unsupported or unknown details callers should not treat as validation proof.</summary>
        public IReadOnlyList<string> UnsupportedDetails { get; }

        /// <summary>Gets human-readable package details suitable for feature reports.</summary>
        public IReadOnlyList<string> Details { get; }
    }

    /// <summary>
    /// Describes one XML signature part in an Open XML package.
    /// </summary>
    internal sealed class OfficePackageSignaturePartInfo {
        internal OfficePackageSignaturePartInfo(
            string uri,
            string contentType,
            string? relationshipId,
            long? length,
            string? signatureMethodAlgorithm,
            IReadOnlyList<string> digestMethodAlgorithms,
            IReadOnlyList<OfficePackageSignatureReferenceInfo> signedReferences,
            IReadOnlyList<OfficePackageSignatureTimestampInfo> timestamps,
            IReadOnlyList<string> x509SubjectNames,
            string? parseError,
            IReadOnlyList<string> unsupportedDetails) {
            Uri = uri;
            ContentType = contentType;
            RelationshipId = relationshipId;
            Length = length;
            SignatureMethodAlgorithm = signatureMethodAlgorithm;
            DigestMethodAlgorithms = digestMethodAlgorithms;
            SignedReferences = signedReferences;
            Timestamps = timestamps;
            X509SubjectNames = x509SubjectNames;
            ParseError = parseError;
            UnsupportedDetails = unsupportedDetails;
        }

        /// <summary>Gets the signature part URI.</summary>
        public string Uri { get; }

        /// <summary>Gets the signature part content type.</summary>
        public string ContentType { get; }

        /// <summary>Gets the relationship id from the signature origin part when available.</summary>
        public string? RelationshipId { get; }

        /// <summary>Gets the signature part byte length when the stream supports length.</summary>
        public long? Length { get; }

        /// <summary>Gets the XML DSig signature method algorithm when parseable.</summary>
        public string? SignatureMethodAlgorithm { get; }

        /// <summary>Gets XML DSig digest method algorithms when parseable.</summary>
        public IReadOnlyList<string> DigestMethodAlgorithms { get; }

        /// <summary>Gets XML DSig signed references discovered in the signature part.</summary>
        public IReadOnlyList<OfficePackageSignatureReferenceInfo> SignedReferences { get; }

        /// <summary>Gets timestamp declarations discovered in the signature XML.</summary>
        public IReadOnlyList<OfficePackageSignatureTimestampInfo> Timestamps { get; }

        /// <summary>Gets XML DSig X509 subject names when parseable.</summary>
        public IReadOnlyList<string> X509SubjectNames { get; }

        /// <summary>Gets the XML parse error, if the signature part could not be parsed.</summary>
        public string? ParseError { get; }

        /// <summary>Gets whether the XML signature part could not be parsed.</summary>
        public bool HasParseError => !string.IsNullOrWhiteSpace(ParseError);

        /// <summary>Gets unsupported or parse details for this signature part.</summary>
        public IReadOnlyList<string> UnsupportedDetails { get; }
    }

    /// <summary>
    /// Describes one XML DSig reference entry in a signature part.
    /// </summary>
    internal sealed class OfficePackageSignatureReferenceInfo {
        internal OfficePackageSignatureReferenceInfo(
            string? uri,
            string? digestMethodAlgorithm,
            string? digestValue,
            bool isPackagePartReference,
            string? targetPartUri,
            bool? targetPartExists,
            IReadOnlyList<string> transformAlgorithms,
            OfficePackageSignatureDigestVerificationStatus digestVerificationStatus,
            string? digestVerificationDetail) {
            Uri = uri;
            DigestMethodAlgorithm = digestMethodAlgorithm;
            DigestValue = digestValue;
            IsPackagePartReference = isPackagePartReference;
            TargetPartUri = targetPartUri;
            TargetPartExists = targetPartExists;
            TransformAlgorithms = transformAlgorithms;
            DigestVerificationStatus = digestVerificationStatus;
            DigestVerificationDetail = digestVerificationDetail;
        }

        /// <summary>Gets the XML DSig Reference URI value.</summary>
        public string? Uri { get; }

        /// <summary>Gets the reference digest method algorithm when parseable.</summary>
        public string? DigestMethodAlgorithm { get; }

        /// <summary>Gets the reference digest value when parseable.</summary>
        public string? DigestValue { get; }

        /// <summary>Gets whether the reference includes a digest value.</summary>
        public bool HasDigestValue => !string.IsNullOrWhiteSpace(DigestValue);

        /// <summary>Gets whether the reference points at an OPC package part URI.</summary>
        public bool IsPackagePartReference { get; }

        /// <summary>Gets the normalized target package part URI when the reference points at a package part.</summary>
        public string? TargetPartUri { get; }

        /// <summary>Gets whether the target package part exists, or null when the reference is not a package part reference.</summary>
        public bool? TargetPartExists { get; }

        /// <summary>Gets XML DSig transform algorithms declared on the reference.</summary>
        public IReadOnlyList<string> TransformAlgorithms { get; }

        /// <summary>Gets bounded digest-verification status after applying supported OPC transforms.</summary>
        public OfficePackageSignatureDigestVerificationStatus DigestVerificationStatus { get; }

        /// <summary>Gets a deterministic digest-verification detail or unsupported reason.</summary>
        public string? DigestVerificationDetail { get; }
    }

    /// <summary>
    /// Describes timestamp metadata declared inside one XML signature part.
    /// </summary>
    internal sealed class OfficePackageSignatureTimestampInfo {
        internal OfficePackageSignatureTimestampInfo(string kind, string? value, string? format) {
            Kind = kind;
            Value = value;
            Format = format;
        }

        /// <summary>Gets the recognized timestamp declaration kind.</summary>
        public string Kind { get; }

        /// <summary>Gets the timestamp value when the declaration exposes one as text.</summary>
        public string? Value { get; }

        /// <summary>Gets the timestamp format when the declaration exposes one.</summary>
        public string? Format { get; }
    }

    /// <summary>
    /// Inspects Open Packaging Convention signature metadata without performing cryptographic validation.
    /// </summary>
    internal static class OfficePackageSignatureInspector {
        internal static OfficePackageSignatureInfo Inspect(
            OpenXmlPackage package,
            DigitalSignatureOriginPart? originPart,
            bool hasApplicationSignatureMetadata,
            byte[]? packageBytes = null,
            int maxPackageParts = 10000,
            long maxPartBytes = 256L * 1024 * 1024,
            int maxSignedReferences = 4096,
            long maxTotalDigestBytes = 512L * 1024 * 1024,
            long maxSignatureBytes = 16L * 1024 * 1024,
            int maxCertificates = 64,
            long maxCertificateBytes = 4L * 1024 * 1024) {
            if (package == null) throw new ArgumentNullException(nameof(package));
            if (maxCertificates <= 0) throw new ArgumentOutOfRangeException(nameof(maxCertificates));
            if (maxCertificateBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maxCertificateBytes));
            if (maxSignedReferences <= 0) throw new ArgumentOutOfRangeException(nameof(maxSignedReferences));
            if (maxTotalDigestBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maxTotalDigestBytes));

            var signatureParts = new List<OfficePackageSignaturePartInfo>();
            var unsupportedDetails = new List<string>();
            var details = new List<string>();
            string? originRelationshipId = null;

            OfficePackageSignatureArchive? signatureArchive = null;
            string? digestInspectionUnavailableDetail = null;
            if (packageBytes != null) {
                try {
                    signatureArchive = new OfficePackageSignatureArchive(packageBytes, maxPackageParts);
                } catch (Exception ex) when (ex is IOException || ex is InvalidDataException) {
                    digestInspectionUnavailableDetail = "Digest inspection was not performed because the bounded OPC archive could not be opened: " + ex.Message;
                    unsupportedDetails.Add(digestInspectionUnavailableDetail);
                }
            }

            try {
                if (originPart != null) {
                    originRelationshipId = FindRelationshipId(package.Parts, originPart);
                    string originDetail = "Digital signature origin part: " + originPart.Uri;
                    if (!string.IsNullOrWhiteSpace(originRelationshipId)) {
                        originDetail += " (" + originRelationshipId + ")";
                    }

                    details.Add(originDetail + ".");

                    Dictionary<string, OpenXmlPart> packageParts = GetPackageParts(package);
                    HashSet<string> packagePartUris = GetPackagePartUris(packageParts);
                    if (signatureArchive != null) packagePartUris.UnionWith(signatureArchive.PartUris);
                    foreach (XmlSignaturePart signaturePart in originPart.XmlSignatureParts) {
                        OfficePackageSignaturePartInfo partInfo = InspectSignaturePart(
                            originPart,
                            signaturePart,
                            packagePartUris,
                            packageParts,
                            signatureArchive,
                            digestInspectionUnavailableDetail,
                            maxPartBytes,
                            maxSignedReferences,
                            maxTotalDigestBytes,
                            maxSignatureBytes,
                            maxCertificates,
                            maxCertificateBytes);
                        signatureParts.Add(partInfo);
                        details.Add(DescribeSignaturePart(partInfo));
                        AddParseDetails(details, "Signature method", partInfo.SignatureMethodAlgorithm);
                        AddParseDetails(details, "Digest methods", partInfo.DigestMethodAlgorithms);
                        AddReferenceDetails(details, partInfo.SignedReferences);
                        AddTimestampDetails(details, partInfo.Timestamps);
                        AddParseDetails(details, "X509 subjects", partInfo.X509SubjectNames);
                        unsupportedDetails.AddRange(partInfo.UnsupportedDetails);
                    }
                }

                if (hasApplicationSignatureMetadata) {
                    details.Add("Extended application properties contain digital signature metadata.");
                }

                if (originPart != null || hasApplicationSignatureMetadata || signatureParts.Count > 0) {
                    details.Add("Signature metadata was inspected independently of caller trust policy.");
                }

                return new OfficePackageSignatureInfo(
                    originPart != null,
                    originPart?.Uri.ToString(),
                    originRelationshipId,
                    hasApplicationSignatureMetadata,
                    signatureParts,
                    unsupportedDetails.Distinct(StringComparer.OrdinalIgnoreCase).ToArray(),
                    details.Distinct(StringComparer.OrdinalIgnoreCase).ToArray());
            } finally {
                signatureArchive?.Dispose();
            }
        }

        private static OfficePackageSignaturePartInfo InspectSignaturePart(
            DigitalSignatureOriginPart originPart,
            XmlSignaturePart signaturePart,
            HashSet<string> packagePartUris,
            IReadOnlyDictionary<string, OpenXmlPart> packageParts,
            OfficePackageSignatureArchive? signatureArchive,
            string? digestInspectionUnavailableDetail,
            long maxPartBytes,
            int maxSignedReferences,
            long maxTotalDigestBytes,
            long maxSignatureBytes,
            int maxCertificates,
            long maxCertificateBytes) {
            var unsupportedDetails = new List<string>();
            string? relationshipId = FindRelationshipId(originPart.Parts, signaturePart);
            long? length = null;
            string? signatureMethod = null;
            string? parseError = null;
            var digestMethods = new List<string>();
            var signedReferences = new List<OfficePackageSignatureReferenceInfo>();
            var timestamps = new List<OfficePackageSignatureTimestampInfo>();
            var subjectNames = new List<string>();

            try {
                using Stream stream = signaturePart.GetStream(FileMode.Open, FileAccess.Read);
                if (stream.CanSeek) {
                    length = stream.Length;
                    if (length.Value > maxSignatureBytes) {
                        throw new InvalidDataException("The XML signature part exceeds the " + maxSignatureBytes + " byte inspection limit.");
                    }
                }

                using var reader = System.Xml.XmlReader.Create(stream, new System.Xml.XmlReaderSettings {
                    DtdProcessing = System.Xml.DtdProcessing.Prohibit,
                    XmlResolver = null,
                    MaxCharactersInDocument = maxSignatureBytes
                });
                XDocument xml = XDocument.Load(reader, LoadOptions.PreserveWhitespace);
                XNamespace ds = "http://www.w3.org/2000/09/xmldsig#";
                signatureMethod = xml.Descendants(ds + "SignatureMethod")
                    .Select(element => (string?)element.Attribute("Algorithm"))
                    .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));
                digestMethods.AddRange(xml.Descendants(ds + "DigestMethod")
                    .Select(element => (string?)element.Attribute("Algorithm"))
                    .Where(value => !string.IsNullOrWhiteSpace(value))
                    .Select(value => value!)
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(value => value, StringComparer.OrdinalIgnoreCase));
                IReadOnlyList<XElement> referencesToInspect = GetAuthenticatedReferences(xml, ds, unsupportedDetails);
                ValidateDigestWorkBudget(referencesToInspect, signatureArchive, maxSignedReferences, maxTotalDigestBytes);
                signedReferences.AddRange(referencesToInspect
                    .Select(reference => InspectSignedReference(reference, ds, packagePartUris, packageParts, signatureArchive, digestInspectionUnavailableDetail, maxPartBytes)));
                timestamps.AddRange(ReadSignatureTimestamps(xml));
                unsupportedDetails.AddRange(signedReferences
                    .Where(reference => reference.DigestVerificationStatus == OfficePackageSignatureDigestVerificationStatus.Unsupported)
                    .Select(reference => reference.DigestVerificationDetail)
                    .Where(detail => !string.IsNullOrWhiteSpace(detail))
                    .Select(detail => detail!));
                subjectNames.AddRange(xml.Descendants(ds + "X509SubjectName")
                    .Select(element => element.Value.Trim())
                    .Where(value => value.Length > 0)
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(value => value, StringComparer.OrdinalIgnoreCase));
                int embeddedCertificateCount = xml.Descendants(ds + "X509Certificate").Count();
                int relatedCertificateCount = signaturePart.Parts.Count(relationship => IsSignatureCertificatePart(relationship.OpenXmlPart));
                if (embeddedCertificateCount + relatedCertificateCount > maxCertificates) {
                    throw new InvalidDataException("The XML signature exceeds the " + maxCertificates + " certificate limit.");
                }
                subjectNames.AddRange(ReadEmbeddedCertificateSubjects(xml, ds, signaturePart.Uri.ToString(), maxCertificateBytes, unsupportedDetails));
                subjectNames.AddRange(ReadRelatedCertificateSubjects(signaturePart, maxCertificateBytes, unsupportedDetails));
                timestamps = timestamps
                    .OrderBy(timestamp => timestamp.Kind, StringComparer.OrdinalIgnoreCase)
                    .ThenBy(timestamp => timestamp.Value, StringComparer.OrdinalIgnoreCase)
                    .ToList();
                subjectNames = subjectNames
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(value => value, StringComparer.OrdinalIgnoreCase)
                    .ToList();
            } catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException || ex is System.Xml.XmlException || ex is InvalidOperationException || ex is InvalidDataException) {
                parseError = ex.Message;
                unsupportedDetails.Add("Unable to parse XML signature part " + signaturePart.Uri + ": " + ex.Message);
            }

            return new OfficePackageSignaturePartInfo(
                signaturePart.Uri.ToString(),
                signaturePart.ContentType,
                relationshipId,
                length,
                signatureMethod,
                digestMethods.ToArray(),
                signedReferences.ToArray(),
                timestamps.ToArray(),
                subjectNames.ToArray(),
                parseError,
                unsupportedDetails.ToArray());
        }

        private static IReadOnlyList<XElement> GetAuthenticatedReferences(
            XDocument xml,
            XNamespace ds,
            List<string> unsupportedDetails) {
            XElement? signedInfo = xml.Root?.Element(ds + "SignedInfo");
            List<XElement> signedInfoReferences = signedInfo?.Elements(ds + "Reference").ToList()
                ?? new List<XElement>();
            var authenticatedManifests = new HashSet<XElement>();
            XElement[] allManifests = xml.Descendants(ds + "Manifest").ToArray();

            foreach (XElement signedReference in signedInfoReferences) {
                string uri = ((string?)signedReference.Attribute("URI"))?.Trim() ?? string.Empty;
                if (uri.Length == 0) {
                    continue;
                }
                if (!uri.StartsWith("#", StringComparison.Ordinal) || uri.Length == 1) continue;

                string id = uri.Substring(1);
                XElement[] targets = xml.Root?
                    .DescendantsAndSelf()
                    .Where(element => HasElementId(element, id))
                    .Take(2)
                    .ToArray() ?? Array.Empty<XElement>();
                if (targets.Length != 1) continue;
                authenticatedManifests.UnionWith(targets[0]
                    .DescendantsAndSelf()
                    .Where(element => element.Name == ds + "Manifest"));
            }

            if (allManifests.Any(manifest => !authenticatedManifests.Contains(manifest))) {
                unsupportedDetails.Add("Ignored package references from an XML DSig Manifest that is not authenticated by SignedInfo.");
            }

            List<XElement> result = signedInfoReferences
                .Where(reference => NormalizePackagePartReference((string?)reference.Attribute("URI")) != null)
                .Concat(authenticatedManifests.SelectMany(manifest => manifest.Elements(ds + "Reference")))
                .Distinct()
                .ToList();
            return result.Count > 0 ? result : signedInfoReferences;
        }

        private static void ValidateDigestWorkBudget(
            IReadOnlyList<XElement> references,
            OfficePackageSignatureArchive? archive,
            int maxSignedReferences,
            long maxTotalDigestBytes) {
            if (references.Count > maxSignedReferences) {
                throw new InvalidDataException("The XML signature contains more than " + maxSignedReferences + " authenticated references.");
            }
            if (archive == null) return;

            long totalDigestBytes = 0;
            foreach (XElement reference in references) {
                string? targetPartUri = NormalizePackagePartReference((string?)reference.Attribute("URI"));
                if (targetPartUri == null || !archive.TryGetPartLength(targetPartUri, out long partLength)) continue;
                if (partLength > maxTotalDigestBytes - totalDigestBytes) {
                    throw new InvalidDataException("Authenticated package references exceed the " + maxTotalDigestBytes + " byte aggregate digest-work limit.");
                }
                totalDigestBytes += partLength;
            }
        }

        private static bool HasElementId(XElement element, string id) {
            foreach (XAttribute attribute in element.Attributes()) {
                bool isId = attribute.Name.Namespace == XNamespace.None &&
                            (attribute.Name.LocalName == "Id" || attribute.Name.LocalName == "ID" || attribute.Name.LocalName == "id");
                if (isId && string.Equals(attribute.Value, id, StringComparison.Ordinal)) return true;
            }
            return false;
        }

        private static IReadOnlyList<OfficePackageSignatureTimestampInfo> ReadSignatureTimestamps(XDocument xml) {
            var timestamps = new List<OfficePackageSignatureTimestampInfo>();

            foreach (XElement signatureTime in xml.Descendants().Where(element =>
                element.Name.LocalName.Equals("SignatureTime", StringComparison.OrdinalIgnoreCase))) {
                string? value = FindDescendantValue(signatureTime, "Value");
                string? format = FindDescendantValue(signatureTime, "Format");
                if (!string.IsNullOrWhiteSpace(value) || !string.IsNullOrWhiteSpace(format)) {
                    timestamps.Add(new OfficePackageSignatureTimestampInfo("OPC SignatureTime", value, format));
                }
            }

            foreach (XElement signingTime in xml.Descendants().Where(element =>
                element.Name.LocalName.Equals("SigningTime", StringComparison.OrdinalIgnoreCase))) {
                string? value = NormalizeText(signingTime.Value);
                if (!string.IsNullOrWhiteSpace(value)) {
                    timestamps.Add(new OfficePackageSignatureTimestampInfo("XAdES SigningTime", value, null));
                }
            }

            return timestamps
                .GroupBy(timestamp => (timestamp.Kind + "\u001f" + timestamp.Value + "\u001f" + timestamp.Format), StringComparer.OrdinalIgnoreCase)
                .Select(group => group.First())
                .OrderBy(timestamp => timestamp.Kind, StringComparer.OrdinalIgnoreCase)
                .ThenBy(timestamp => timestamp.Value, StringComparer.OrdinalIgnoreCase)
                .ToArray();
        }

        private static string? FindDescendantValue(XElement element, string localName) {
            return element
                .Descendants()
                .Where(descendant => descendant.Name.LocalName.Equals(localName, StringComparison.OrdinalIgnoreCase))
                .Select(descendant => NormalizeText(descendant.Value))
                .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));
        }

        private static string? NormalizeText(string? value) {
            if (string.IsNullOrWhiteSpace(value)) {
                return null;
            }

            return value!.Trim();
        }

        private static IEnumerable<string> ReadEmbeddedCertificateSubjects(
            XDocument xml,
            XNamespace ds,
            string signaturePartUri,
            long maxCertificateBytes,
            List<string> unsupportedDetails) {
            foreach (XElement element in xml.Descendants(ds + "X509Certificate")) {
                string certificateText = element.Value.Trim();
                if (certificateText.Length == 0) {
                    continue;
                }

                byte[] rawCertificate;
                try {
                    if (certificateText.Length > GetMaxBase64EncodedCharacters(maxCertificateBytes)) {
                        throw new InvalidDataException("The embedded X509Certificate exceeds the " + maxCertificateBytes + " byte limit.");
                    }
                    rawCertificate = Convert.FromBase64String(certificateText);
                    if (rawCertificate.LongLength > maxCertificateBytes) {
                        throw new InvalidDataException("The embedded X509Certificate exceeds the " + maxCertificateBytes + " byte limit.");
                    }
                } catch (FormatException ex) {
                    unsupportedDetails.Add("Unable to parse X509Certificate in XML signature part " + signaturePartUri + ": " + ex.Message);
                    continue;
                }

                string? subject = ReadCertificateSubject(rawCertificate, "embedded X509Certificate in XML signature part " + signaturePartUri, unsupportedDetails);
                if (!string.IsNullOrWhiteSpace(subject)) {
                    yield return subject!;
                }
            }
        }

        private static IEnumerable<string> ReadRelatedCertificateSubjects(
            XmlSignaturePart signaturePart,
            long maxCertificateBytes,
            List<string> unsupportedDetails) {
            foreach (IdPartPair relationship in signaturePart.Parts) {
                OpenXmlPart relatedPart = relationship.OpenXmlPart;
                if (!IsSignatureCertificatePart(relatedPart)) {
                    continue;
                }

                byte[] rawCertificate;
                try {
                    using Stream stream = relatedPart.GetStream(FileMode.Open, FileAccess.Read);
                    if (stream.CanSeek && stream.Length > maxCertificateBytes) {
                        throw new InvalidDataException("The signature certificate part exceeds the " + maxCertificateBytes + " byte limit.");
                    }
                    using var memoryStream = new MemoryStream();
                    CopyBounded(stream, memoryStream, maxCertificateBytes);
                    rawCertificate = memoryStream.ToArray();
                } catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException || ex is InvalidOperationException) {
                    unsupportedDetails.Add("Unable to read signature certificate part " + relatedPart.Uri + ": " + ex.Message);
                    continue;
                }

                string? subject = ReadCertificateSubject(rawCertificate, "signature certificate part " + relatedPart.Uri, unsupportedDetails);
                if (!string.IsNullOrWhiteSpace(subject)) {
                    yield return subject!;
                }
            }
        }

        private static bool IsSignatureCertificatePart(OpenXmlPart part) {
            return part.RelationshipType.EndsWith("/digital-signature/certificate", StringComparison.OrdinalIgnoreCase) ||
                   part.Uri.ToString().EndsWith(".cer", StringComparison.OrdinalIgnoreCase);
        }

        private static string? ReadCertificateSubject(byte[] rawCertificate, string source, List<string> unsupportedDetails) {
            try {
                using X509Certificate2 certificate = LoadCertificate(rawCertificate);
                string subjectName = certificate.SubjectName.Name ?? certificate.Subject;
                if (!string.IsNullOrWhiteSpace(subjectName)) {
                    return subjectName.Trim();
                }
            } catch (CryptographicException ex) {
                unsupportedDetails.Add("Unable to parse X509 certificate from " + source + ": " + ex.Message);
            }

            return null;
        }

        private static X509Certificate2 LoadCertificate(byte[] rawCertificate) {
#if NET9_0_OR_GREATER
            return X509CertificateLoader.LoadCertificate(rawCertificate);
#else
            return new X509Certificate2(rawCertificate);
#endif
        }

        private static long GetMaxBase64EncodedCharacters(long maxDecodedBytes) {
            return maxDecodedBytes > (long.MaxValue / 4L) * 3L
                ? long.MaxValue
                : ((maxDecodedBytes + 2L) / 3L) * 4L;
        }

        private static void CopyBounded(Stream source, Stream destination, long maxBytes) {
            byte[] buffer = new byte[81920];
            long total = 0;
            int read;
            while ((read = source.Read(buffer, 0, buffer.Length)) > 0) {
                total += read;
                if (total > maxBytes) {
                    throw new InvalidDataException("The signature certificate part exceeds the " + maxBytes + " byte limit.");
                }
                destination.Write(buffer, 0, read);
            }
        }

        private static OfficePackageSignatureReferenceInfo InspectSignedReference(
            XElement reference,
            XNamespace ds,
            HashSet<string> packagePartUris,
            IReadOnlyDictionary<string, OpenXmlPart> packageParts,
            OfficePackageSignatureArchive? signatureArchive,
            string? digestInspectionUnavailableDetail,
            long maxPartBytes) {
            string? uri = ((string?)reference.Attribute("URI"))?.Trim();
            string? digestMethod = reference.Element(ds + "DigestMethod")?.Attribute("Algorithm")?.Value;
            string? digestValue = reference.Element(ds + "DigestValue")?.Value.Trim();
            string? targetPartUri = NormalizePackagePartReference(uri);
            bool? targetPartExists = targetPartUri == null ? null : packagePartUris.Contains(targetPartUri);
            string[] transformAlgorithms = reference
                .Descendants(ds + "Transform")
                .Select(element => (string?)element.Attribute("Algorithm"))
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Select(value => value!)
                .ToArray();
            DigestVerificationResult digestVerification = VerifyReferenceDigest(
                targetPartUri,
                targetPartExists,
                digestMethod,
                digestValue,
                transformAlgorithms,
                packageParts,
                reference,
                signatureArchive,
                digestInspectionUnavailableDetail,
                maxPartBytes);

            return new OfficePackageSignatureReferenceInfo(
                uri,
                string.IsNullOrWhiteSpace(digestMethod) ? null : digestMethod,
                string.IsNullOrWhiteSpace(digestValue) ? null : digestValue,
                targetPartUri != null,
                targetPartUri,
                targetPartExists,
                transformAlgorithms,
                digestVerification.Status,
                digestVerification.Detail);
        }

        private static DigestVerificationResult VerifyReferenceDigest(
            string? targetPartUri,
            bool? targetPartExists,
            string? digestMethod,
            string? digestValue,
            IReadOnlyList<string> transformAlgorithms,
            IReadOnlyDictionary<string, OpenXmlPart> packageParts,
            XElement reference,
            OfficePackageSignatureArchive? signatureArchive,
            string? digestInspectionUnavailableDetail,
            long maxPartBytes) {
            if (string.IsNullOrWhiteSpace(targetPartUri) || targetPartExists != true) {
                return DigestVerificationResult.NotChecked(null);
            }

            if (string.IsNullOrWhiteSpace(digestMethod) || string.IsNullOrWhiteSpace(digestValue)) {
                return DigestVerificationResult.NotChecked(null);
            }

            string normalizedTargetPartUri = targetPartUri!;
            string normalizedDigestMethod = digestMethod!;
            string normalizedDigestValue = digestValue!;

            if (!string.IsNullOrWhiteSpace(digestInspectionUnavailableDetail)) {
                return DigestVerificationResult.Unsupported(digestInspectionUnavailableDetail!);
            }

            if (signatureArchive != null) {
                OfficePackageDigestResult transformed = signatureArchive.VerifyReference(reference, maxPartBytes);
                switch (transformed.Status) {
                    case OfficePackageSignatureDigestVerificationStatus.Passed:
                        return DigestVerificationResult.Passed(transformed.Detail);
                    case OfficePackageSignatureDigestVerificationStatus.Failed:
                        return DigestVerificationResult.Failed(transformed.Detail);
                    case OfficePackageSignatureDigestVerificationStatus.Unsupported:
                        return DigestVerificationResult.Unsupported(transformed.Detail);
                    default:
                        return DigestVerificationResult.NotChecked(transformed.Detail);
                }
            }

            if (transformAlgorithms.Count > 0) {
                return DigestVerificationResult.Unsupported("Digest verification for " + normalizedTargetPartUri + " was not checked because the reference declares XML DSig transforms.");
            }

            if (!packageParts.TryGetValue(normalizedTargetPartUri, out OpenXmlPart? part)) {
                return DigestVerificationResult.Unsupported("Digest verification for " + normalizedTargetPartUri + " was not checked because the target is not a directly readable package part.");
            }

            Func<HashAlgorithm>? hashFactory = CreateHashAlgorithm(normalizedDigestMethod);
            if (hashFactory == null) {
                return DigestVerificationResult.Unsupported("Digest verification for " + normalizedTargetPartUri + " was not checked because digest method " + normalizedDigestMethod + " is not supported.");
            }

            byte[] expectedDigest;
            try {
                expectedDigest = Convert.FromBase64String(normalizedDigestValue);
            } catch (FormatException ex) {
                return DigestVerificationResult.Failed("Digest verification for " + normalizedTargetPartUri + " failed because DigestValue is not valid base64: " + ex.Message);
            }

            byte[] actualDigest;
            try {
                using Stream stream = part.GetStream(FileMode.Open, FileAccess.Read);
                if (stream.CanSeek && stream.Length > maxPartBytes) {
                    return DigestVerificationResult.Unsupported("Digest verification for " + normalizedTargetPartUri + " was not checked because the package part exceeds the " + maxPartBytes + " byte limit.");
                }
                using HashAlgorithm hashAlgorithm = hashFactory();
                actualDigest = hashAlgorithm.ComputeHash(stream);
            } catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException || ex is InvalidOperationException || ex is CryptographicException) {
                return DigestVerificationResult.Unsupported("Digest verification for " + normalizedTargetPartUri + " was not checked because the package part could not be read: " + ex.Message);
            }

            if (actualDigest.SequenceEqual(expectedDigest)) {
                return DigestVerificationResult.Passed("Digest verification passed for " + normalizedTargetPartUri + ".");
            }

            return DigestVerificationResult.Failed("Digest verification failed for " + normalizedTargetPartUri + ".");
        }

        private static Func<HashAlgorithm>? CreateHashAlgorithm(string digestMethod) {
            switch (digestMethod.Trim()) {
                case "http://www.w3.org/2000/09/xmldsig#sha1":
                    return SHA1.Create;
                case "http://www.w3.org/2001/04/xmlenc#sha256":
                case "http://www.w3.org/2001/04/xmldsig-more#sha256":
                    return SHA256.Create;
                case "http://www.w3.org/2001/04/xmldsig-more#sha384":
                    return SHA384.Create;
                case "http://www.w3.org/2001/04/xmlenc#sha512":
                case "http://www.w3.org/2001/04/xmldsig-more#sha512":
                    return SHA512.Create;
                default:
                    return null;
            }
        }

        private static string? NormalizePackagePartReference(string? uri) {
            if (string.IsNullOrWhiteSpace(uri)) {
                return null;
            }

            string trimmed = uri!.Trim();
            if (trimmed.StartsWith("#", StringComparison.Ordinal)) {
                return null;
            }

            if (!trimmed.StartsWith("/", StringComparison.Ordinal)) {
                if (Uri.TryCreate(trimmed, UriKind.Absolute, out Uri? absoluteUri) && !string.IsNullOrWhiteSpace(absoluteUri.Scheme)) {
                    return null;
                }

                return null;
            }

            int fragmentIndex = trimmed.IndexOf('#');
            if (fragmentIndex >= 0) {
                return null;
            }

            int queryIndex = trimmed.IndexOf('?');
            if (queryIndex >= 0) {
                trimmed = trimmed.Substring(0, queryIndex);
            }

            return trimmed.Length == 0 ? null : trimmed;
        }

        private static string DescribeSignaturePart(OfficePackageSignaturePartInfo partInfo) {
            string detail = "XML signature part: " + partInfo.Uri;
            if (!string.IsNullOrWhiteSpace(partInfo.RelationshipId)) {
                detail += " (" + partInfo.RelationshipId + ")";
            }

            if (partInfo.Length.HasValue) {
                detail += ", " + partInfo.Length.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) + " bytes";
            }

            return detail + ".";
        }

        private static void AddParseDetails(List<string> details, string label, string? value) {
            if (!string.IsNullOrWhiteSpace(value)) {
                details.Add(label + ": " + value + ".");
            }
        }

        private static void AddParseDetails(List<string> details, string label, IReadOnlyList<string> values) {
            if (values.Count > 0) {
                details.Add(label + ": " + string.Join(", ", values) + ".");
            }
        }

        private static void AddReferenceDetails(List<string> details, IReadOnlyList<OfficePackageSignatureReferenceInfo> references) {
            foreach (OfficePackageSignatureReferenceInfo reference in references) {
                string referenceUri = string.IsNullOrWhiteSpace(reference.Uri) ? "(empty)" : reference.Uri!;
                string detail = "Signed reference: " + referenceUri;
                if (!string.IsNullOrWhiteSpace(reference.DigestMethodAlgorithm)) {
                    detail += " (" + reference.DigestMethodAlgorithm + ")";
                }

                detail += reference.HasDigestValue ? " with digest value" : " without digest value";

                if (reference.IsPackagePartReference) {
                    detail += reference.TargetPartExists == true ? " targets an existing package part" : " targets a missing package part";
                } else {
                    detail += " is not a package part reference";
                }

                details.Add(detail + ".");
                if (!string.IsNullOrWhiteSpace(reference.DigestVerificationDetail)) {
                    details.Add(reference.DigestVerificationDetail!);
                }
            }
        }

        private static void AddTimestampDetails(List<string> details, IReadOnlyList<OfficePackageSignatureTimestampInfo> timestamps) {
            foreach (OfficePackageSignatureTimestampInfo timestamp in timestamps) {
                string detail = "Signature timestamp: " + timestamp.Kind;
                if (!string.IsNullOrWhiteSpace(timestamp.Value)) {
                    detail += " value " + timestamp.Value;
                }

                if (!string.IsNullOrWhiteSpace(timestamp.Format)) {
                    detail += " (" + timestamp.Format + ")";
                }

                details.Add(detail + ".");
            }
        }

        private static Dictionary<string, OpenXmlPart> GetPackageParts(OpenXmlPackage package) {
            var parts = new Dictionary<string, OpenXmlPart>(StringComparer.OrdinalIgnoreCase);
            foreach (IdPartPair pair in package.Parts) {
                AddPackageParts(pair.OpenXmlPart, parts);
            }

            return parts;
        }

        private static void AddPackageParts(OpenXmlPart part, Dictionary<string, OpenXmlPart> parts) {
            string partUri = part.Uri.ToString();
            if (parts.ContainsKey(partUri)) {
                return;
            }

            parts.Add(partUri, part);
            foreach (IdPartPair child in part.Parts) {
                AddPackageParts(child.OpenXmlPart, parts);
            }
        }

        private static HashSet<string> GetPackagePartUris(IReadOnlyDictionary<string, OpenXmlPart> packageParts) {
            var uris = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (packageParts.Count > 0) {
                uris.Add("/_rels/.rels");
            }

            foreach (OpenXmlPart part in packageParts.Values) {
                AddPackagePartUris(part, uris);
            }

            return uris;
        }

        private static void AddPackagePartUris(OpenXmlPart part, HashSet<string> uris) {
            if (!uris.Add(part.Uri.ToString())) {
                return;
            }

            if (HasRelationships(part)) {
                uris.Add(GetRelationshipPartUri(part.Uri));
            }

            foreach (IdPartPair child in part.Parts) {
                AddPackagePartUris(child.OpenXmlPart, uris);
            }
        }

        private static bool HasRelationships(OpenXmlPart part) {
            return part.Parts.Any() ||
                   part.ExternalRelationships.Any() ||
                   part.HyperlinkRelationships.Any() ||
                   part.DataPartReferenceRelationships.Any();
        }

        private static string GetRelationshipPartUri(Uri partUri) {
            string partPath = partUri.ToString();
            int slashIndex = partPath.LastIndexOf('/');
            if (slashIndex < 0) {
                return "/_rels/" + partPath + ".rels";
            }

            string folder = partPath.Substring(0, slashIndex + 1);
            string fileName = partPath.Substring(slashIndex + 1);
            return folder + "_rels/" + fileName + ".rels";
        }

        private static string? FindRelationshipId(IEnumerable<IdPartPair> pairs, OpenXmlPart part) {
            foreach (IdPartPair pair in pairs) {
                if (ReferenceEquals(pair.OpenXmlPart, part)) {
                    return pair.RelationshipId;
                }
            }

            return null;
        }

        private sealed class DigestVerificationResult {
            private DigestVerificationResult(OfficePackageSignatureDigestVerificationStatus status, string? detail) {
                Status = status;
                Detail = detail;
            }

            internal OfficePackageSignatureDigestVerificationStatus Status { get; }

            internal string? Detail { get; }

            internal static DigestVerificationResult NotChecked(string? detail) {
                return new DigestVerificationResult(OfficePackageSignatureDigestVerificationStatus.NotChecked, detail);
            }

            internal static DigestVerificationResult Passed(string detail) {
                return new DigestVerificationResult(OfficePackageSignatureDigestVerificationStatus.Passed, detail);
            }

            internal static DigestVerificationResult Failed(string detail) {
                return new DigestVerificationResult(OfficePackageSignatureDigestVerificationStatus.Failed, detail);
            }

            internal static DigestVerificationResult Unsupported(string detail) {
                return new DigestVerificationResult(OfficePackageSignatureDigestVerificationStatus.Unsupported, detail);
            }
        }
    }
}
