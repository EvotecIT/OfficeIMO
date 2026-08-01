#nullable enable
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing.Internal;
using System.Diagnostics.CodeAnalysis;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Security.Cryptography.Xml;
using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO.Word {
    /// <summary>Options for signing an Open Packaging Convention package.</summary>
    internal sealed class OfficePackageSigningOptions {
        internal const string Sha256HashAlgorithm = "http://www.w3.org/2001/04/xmlenc#sha256";

        public IReadOnlyCollection<string>? PartUris { get; set; }
        public bool IncludePackageRelationships { get; set; } = true;
        public bool IncludePartRelationships { get; set; } = true;
        public string HashAlgorithm { get; set; } = Sha256HashAlgorithm;
        public string? SignatureId { get; set; }
        public DateTimeOffset? SigningTime { get; set; }
        public IReadOnlyCollection<X509Certificate2>? AdditionalCertificates { get; set; }
        public int MaxPackageParts { get; set; } = 10000;
        public long MaxPackageBytes { get; set; } = 512L * 1024 * 1024;
        public long MaxPartBytes { get; set; } = 256L * 1024 * 1024;
        public long MaxTotalDigestBytes { get; set; } = 512L * 1024 * 1024;
        public int MaxSignedReferences { get; set; } = 4096;
        public long MaxSignatureBytes { get; set; } = 16L * 1024 * 1024;
        public int MaxCertificates { get; set; } = 64;
        public long MaxCertificateBytes { get; set; } = 4L * 1024 * 1024;
        public long MaxTotalCertificateBytes { get; set; } = 64L * 1024 * 1024;
        internal Action<string, string>? BeforeCommit { get; set; }
    }

    /// <summary>Result of an attempted Open Packaging Convention package-signing operation.</summary>
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

    /// <summary>Creates interoperable OPC XML signatures through cross-platform cryptographic primitives.</summary>
    internal static class OfficePackageSignatureWriter {
        private const string DigitalSignatureNamespace = "http://schemas.openxmlformats.org/package/2006/digital-signature";
        private const string ObjectReferenceType = "http://www.w3.org/2000/09/xmldsig#Object";
        private const string PackageObjectId = "idPackageObject";
        private const string SignatureTimePropertyId = "idSignatureTime";

        internal static OfficePackageSigningResult Sign(
            string filePath,
            X509Certificate2 certificate,
            OfficePackageSigningOptions? options = null) {
            options ??= new OfficePackageSigningOptions();
            if (string.IsNullOrWhiteSpace(filePath)) return Failed(filePath ?? string.Empty, "A package path is required.");

            string fullPath = Path.GetFullPath(filePath);
            if (!File.Exists(fullPath)) return Failed(fullPath, "The package file does not exist.");
            if (certificate == null) return Failed(fullPath, "A signing certificate is required.");
            if (options.MaxPackageParts <= 0) return Failed(fullPath, "MaxPackageParts must be greater than zero.");
            if (options.MaxPackageBytes <= 0) return Failed(fullPath, "MaxPackageBytes must be greater than zero.");
            if (options.MaxPartBytes <= 0) return Failed(fullPath, "MaxPartBytes must be greater than zero.");
            if (options.MaxTotalDigestBytes <= 0) return Failed(fullPath, "MaxTotalDigestBytes must be greater than zero.");
            if (options.MaxSignedReferences <= 0) return Failed(fullPath, "MaxSignedReferences must be greater than zero.");
            if (options.MaxSignatureBytes <= 0) return Failed(fullPath, "MaxSignatureBytes must be greater than zero.");
            if (options.MaxCertificates <= 0) return Failed(fullPath, "MaxCertificates must be greater than zero.");
            if (options.MaxCertificateBytes <= 0) return Failed(fullPath, "MaxCertificateBytes must be greater than zero.");
            if (options.MaxTotalCertificateBytes <= 0) return Failed(fullPath, "MaxTotalCertificateBytes must be greater than zero.");
            long packageLength = new FileInfo(fullPath).Length;
            if (packageLength > options.MaxPackageBytes) {
                return Failed(fullPath, "The package exceeds the " + options.MaxPackageBytes + " byte signing limit.");
            }

            string stagingPath = string.Empty;
            try {
                if (!certificate.HasPrivateKey) return Failed(fullPath, "The signing certificate must include a private key.");
                using RSA? signingKey = certificate.GetRSAPrivateKey();
                if (signingKey == null) return Failed(fullPath, "OPC package signing requires an RSA certificate with an accessible private key.");
                IReadOnlyList<X509Certificate2> signingCertificates = ValidateSigningCertificates(certificate, options);

                stagingPath = OfficeFileCommit.CreateStagingPath(fullPath);
                WordPackageSnapshot.CopyBounded(fullPath, stagingPath, options.MaxPackageBytes);
                string sourcePackageHash = WordPackageSnapshot.ComputeSha256(stagingPath, options.MaxPackageBytes);
                PrepareDigitalSignatureMetadata(stagingPath);
                byte[] packageBytes = File.ReadAllBytes(stagingPath);
                SigningPayload payload = CreateSignature(packageBytes, signingCertificates, signingKey, options);
                SignaturePartWriteResult write = AddSignaturePart(stagingPath, payload.SignatureXml);
                EnsureSignedPackageWithinLimits(stagingPath, options);
                string validatedPackageHash = WordPackageSnapshot.ComputeSha256(stagingPath, options.MaxPackageBytes);
                options.BeforeCommit?.Invoke(stagingPath, fullPath);
                if (!string.Equals(
                    validatedPackageHash,
                    WordPackageSnapshot.ComputeSha256(stagingPath, options.MaxPackageBytes),
                    StringComparison.Ordinal)) {
                    return Failed(fullPath, "The validated staging package changed before atomic commit; the original source was preserved.");
                }
                if (!OfficeFileCommit.TryCommitTemporaryFileAtomicallyIfDestinationUnchanged(
                    stagingPath,
                    fullPath,
                    displacedPath => string.Equals(
                        sourcePackageHash,
                        WordPackageSnapshot.ComputeSha256(displacedPath, options.MaxPackageBytes),
                        StringComparison.Ordinal),
                    installedPath => string.Equals(
                        validatedPackageHash,
                        WordPackageSnapshot.ComputeSha256(installedPath, options.MaxPackageBytes),
                        StringComparison.Ordinal))) {
                    return Failed(fullPath, "The source or validated staging package changed while its signature was being created; the current source was preserved.");
                }
                stagingPath = string.Empty;

                string[] details = {
                    "Package signature was created with " + payload.SignedPartCount + " signed package part(s).",
                    "Signed relationship selector count: " + payload.RelationshipSelectorCount + ".",
                    "Signature count after signing: " + write.SignatureCount + ".",
                    "The package signature was written through the cross-platform OPC XML-signature engine."
                };
                return new OfficePackageSigningResult(
                    fullPath,
                    isSupported: true,
                    succeeded: true,
                    payload.SignedPartCount,
                    payload.RelationshipSelectorCount,
                    write.SignatureCount,
                    write.SignaturePartUri,
                    details);
            } catch (Exception exception) when (IsSigningException(exception)) {
                return Failed(fullPath, "Package signing failed: " + exception.Message);
            } finally {
                if (!string.IsNullOrWhiteSpace(stagingPath)) OfficeFileCommit.DeleteIfExists(stagingPath);
            }
        }

        private static void EnsureSignedPackageWithinLimits(string packagePath, OfficePackageSigningOptions options) {
            long packageLength = new FileInfo(packagePath).Length;
            if (packageLength > options.MaxPackageBytes) {
                throw new InvalidDataException("The signed package exceeds the " + options.MaxPackageBytes + " byte signing limit.");
            }
            using var archive = new OfficePackageSignatureArchive(File.ReadAllBytes(packagePath), options.MaxPackageParts);
        }

        private static IReadOnlyList<X509Certificate2> ValidateSigningCertificates(X509Certificate2 signer, OfficePackageSigningOptions options) {
            var thumbprints = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var certificates = new List<X509Certificate2>();
            long totalCertificateBytes = 0;

            ValidateCertificate(signer);
            if (options.AdditionalCertificates != null) {
                foreach (X509Certificate2 additional in options.AdditionalCertificates) {
                    if (additional != null) ValidateCertificate(additional);
                }
            }

            void ValidateCertificate(X509Certificate2 certificate) {
                string identity = GetCertificateIdentity(certificate);
                if (!thumbprints.Add(identity)) return;
                certificates.Add(certificate);
                if (certificates.Count > options.MaxCertificates) {
                    throw new InvalidDataException("The signing certificate set exceeds the " + options.MaxCertificates + " certificate limit.");
                }
                long certificateBytes = certificate.RawData.LongLength;
                if (certificateBytes > options.MaxCertificateBytes) {
                    throw new InvalidDataException("A signing certificate exceeds the " + options.MaxCertificateBytes + " byte limit.");
                }
                if (certificateBytes > options.MaxTotalCertificateBytes - totalCertificateBytes) {
                    throw new InvalidDataException("The signing certificate set exceeds the " + options.MaxTotalCertificateBytes + " byte aggregate certificate limit.");
                }
                totalCertificateBytes += certificateBytes;
            }

            return certificates;
        }

        private static string GetCertificateIdentity(X509Certificate2 certificate) =>
            certificate.Thumbprint ?? Convert.ToBase64String(certificate.RawData);

        private static void PrepareDigitalSignatureMetadata(string packagePath) {
            using WordprocessingDocument package = WordprocessingDocument.Open(packagePath, true);
            if (package.DigitalSignatureOriginPart?.XmlSignatureParts.Any() == true) {
                return;
            }
            ExtendedFilePropertiesPart appPart = package.ExtendedFilePropertiesPart ?? package.AddExtendedFilePropertiesPart();
            appPart.Properties ??= new Properties();
            appPart.Properties.DigitalSignature ??= new DigitalSignature();
            appPart.Properties.Save();
        }

        [UnconditionalSuppressMessage("Trimming", "IL2026", Justification = "Package signing selects a closed RSA and canonicalization algorithm set in code; it does not resolve algorithm implementations from caller-supplied XML.")]
        [UnconditionalSuppressMessage("AOT", "IL3050", Justification = "Package signing never uses the XSLT XML DSig transform and does not compile dynamic XSLT code.")]
        private static SigningPayload CreateSignature(
            byte[] packageBytes,
            IReadOnlyList<X509Certificate2> signingCertificates,
            RSA signingKey,
            OfficePackageSigningOptions options) {
            using var package = new OfficePackageSignatureArchive(packageBytes, options.MaxPackageParts);
            List<string> partUris = ResolvePartUris(package, options, out IReadOnlyList<string> missingPartUris);
            if (missingPartUris.Count > 0) {
                throw new InvalidOperationException("Requested signing part(s) were not found: " + string.Join(", ", missingPartUris) + ".");
            }
            if (partUris.Count == 0) throw new InvalidOperationException("No package parts were selected for signing.");

            string signatureId = ResolveSignatureId(options.SignatureId);
            XNamespace ds = SignedXml.XmlDsigNamespaceUrl;
            XNamespace opc = DigitalSignatureNamespace;
            var manifest = new XElement(ds + "Manifest");
            int authenticatedReferenceCount = checked(partUris.Count + 1);
            if (authenticatedReferenceCount > options.MaxSignedReferences) {
                throw new InvalidDataException("The package signature would contain more than " + options.MaxSignedReferences + " authenticated references.");
            }
            int remainingRelationshipSelectors = options.MaxSignedReferences - authenticatedReferenceCount;
            long totalDigestBytes = 0;
            foreach (string partUri in partUris) {
                if (!package.TryGetContentType(partUri, out string contentType)) {
                    throw new InvalidDataException("The OPC package does not declare a content type for " + partUri + ".");
                }
                ReserveDigestBytes(package, partUri, options, ref totalDigestBytes);
                manifest.Add(CreateReference(partUri + "?ContentType=" + Uri.EscapeDataString(contentType), options.HashAlgorithm));
            }

            int relationshipSelectorCount = 0;
            if (options.IncludePackageRelationships) {
                XElement? packageRelationships = CreateRelationshipReference(
                    package,
                    "/_rels/.rels",
                    options,
                    ref remainingRelationshipSelectors,
                    ref totalDigestBytes,
                    out int selectorCount);
                if (packageRelationships != null) {
                    relationshipSelectorCount += selectorCount;
                    manifest.Add(packageRelationships);
                }
            }
            if (options.IncludePartRelationships) {
                foreach (string partUri in partUris) {
                    string relationshipPartUri = GetRelationshipPartUri(partUri);
                    XElement? partRelationships = CreateRelationshipReference(
                        package,
                        relationshipPartUri,
                        options,
                        ref remainingRelationshipSelectors,
                        ref totalDigestBytes,
                        out int selectorCount);
                    if (partRelationships == null) continue;
                    relationshipSelectorCount += selectorCount;
                    manifest.Add(partRelationships);
                }
            }

            foreach (XElement reference in manifest.Elements(ds + "Reference")) {
                string? referenceUri = ((string?)reference.Attribute("URI"))?.Trim();
                string targetPartUri = OfficePackageSignatureArchive.NormalizeReferencePartUri(referenceUri)
                    ?? throw new InvalidDataException("The OPC signature Reference is not a package-part URI: " + referenceUri + ".");
                if (!package.TryGetPartLength(targetPartUri, out _)) {
                    throw new FileNotFoundException("The package part selected for signing was not found.", targetPartUri);
                }
                reference.Add(new XElement(ds + "DigestValue", package.ComputeDigestValue(reference, options.MaxPartBytes)));
            }

            DateTimeOffset signingTime = options.SigningTime ?? DateTimeOffset.UtcNow;
            var signatureProperties = new XElement(
                ds + "SignatureProperties",
                new XElement(
                    ds + "SignatureProperty",
                    new XAttribute("Id", SignatureTimePropertyId),
                    new XAttribute("Target", "#" + signatureId),
                    new XElement(
                        opc + "SignatureTime",
                        new XElement(opc + "Format", "YYYY-MM-DDThh:mm:ss.sTZD"),
                        new XElement(opc + "Value", signingTime.ToString("yyyy-MM-dd'T'HH:mm:ss.FFFFFFFK", System.Globalization.CultureInfo.InvariantCulture)))));

            XmlDocument objectDocument = CreateXmlDocument();
            objectDocument.LoadXml(new XElement("Root", manifest, signatureProperties).ToString(SaveOptions.DisableFormatting));
            var packageObject = new DataObject {
                Id = PackageObjectId,
                Data = objectDocument.DocumentElement!.ChildNodes
            };

            XmlDocument signatureDocument = CreateXmlDocument();
            var signedXml = new SignedXml(signatureDocument) {
                SigningKey = signingKey
            };
            signedXml.Signature.Id = signatureId;
            signedXml.SignedInfo!.CanonicalizationMethod = SignedXml.XmlDsigC14NTransformUrl;
            signedXml.SignedInfo.SignatureMethod = ResolveSignatureMethod(options.HashAlgorithm);
            signedXml.AddObject(packageObject);
            signedXml.AddReference(new Reference("#" + PackageObjectId) {
                Type = ObjectReferenceType,
                DigestMethod = options.HashAlgorithm
            });

            var keyInfo = new KeyInfo();
            var x509Data = new KeyInfoX509Data(signingCertificates[0]);
            for (int certificateIndex = 1; certificateIndex < signingCertificates.Count; certificateIndex++) {
                x509Data.AddCertificate(signingCertificates[certificateIndex]);
            }
            keyInfo.AddClause(x509Data);
            signedXml.KeyInfo = keyInfo;
            signedXml.ComputeSignature();

            XmlElement signature = signedXml.GetXml();
            signatureDocument.AppendChild(signatureDocument.ImportNode(signature, deep: true));
            using var output = new SignatureBoundedMemoryStream(options.MaxSignatureBytes);
            using (XmlWriter writer = XmlWriter.Create(output, new XmlWriterSettings {
                Encoding = new System.Text.UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
                Indent = false,
                OmitXmlDeclaration = false
            })) {
                signatureDocument.Save(writer);
            }
            return new SigningPayload(output.ToArray(), partUris.Count, relationshipSelectorCount);
        }

        private static SignaturePartWriteResult AddSignaturePart(string packagePath, byte[] signatureXml) {
            string generatedUri;
            int signatureCount;
            using (WordprocessingDocument package = WordprocessingDocument.Open(packagePath, true)) {
                DigitalSignatureOriginPart origin = package.DigitalSignatureOriginPart ?? package.AddDigitalSignatureOriginPart();
                XmlSignaturePart signaturePart = origin.AddNewPart<XmlSignaturePart>();
                using (var stream = new MemoryStream(signatureXml, writable: false)) signaturePart.FeedData(stream);
                generatedUri = signaturePart.Uri.ToString();
                signatureCount = origin.XmlSignatureParts.Count();
            }

            string normalizedUri = NormalizeSignaturePartUri(packagePath, generatedUri);
            return new SignaturePartWriteResult(normalizedUri, signatureCount);
        }

        private static string NormalizeSignaturePartUri(string packagePath, string generatedUri) {
            const string redundantPrefix = "/_xmlsignatures/_xmlsignatures/";
            if (!generatedUri.StartsWith(redundantPrefix, StringComparison.OrdinalIgnoreCase)) return generatedUri;

            using var archive = ZipFile.Open(packagePath, ZipArchiveMode.Update);
            string sourceName = generatedUri.TrimStart('/');
            ZipArchiveEntry source = archive.GetEntry(sourceName)
                ?? throw new InvalidDataException("Generated OPC signature part was not found at " + generatedUri + ".");
            string fileName = Path.GetFileName(sourceName);
            string targetName = "_xmlsignatures/" + fileName;
            for (int suffix = 1; archive.GetEntry(targetName) != null; suffix++) {
                targetName = "_xmlsignatures/sig" + suffix.ToString(System.Globalization.CultureInfo.InvariantCulture) + ".xml";
            }

            byte[] bytes;
            using (Stream input = source.Open()) {
                using var buffer = new MemoryStream();
                input.CopyTo(buffer);
                bytes = buffer.ToArray();
            }
            source.Delete();
            ZipArchiveEntry target = archive.CreateEntry(targetName, CompressionLevel.Optimal);
            using (Stream output = target.Open()) output.Write(bytes, 0, bytes.Length);

            RewriteXmlEntry(archive, "_xmlsignatures/_rels/origin.sigs.rels", document => {
                XNamespace relationships = "http://schemas.openxmlformats.org/package/2006/relationships";
                XElement relationship = document.Descendants(relationships + "Relationship")
                    .Single(element => string.Equals((string?)element.Attribute("Target"), generatedUri, StringComparison.OrdinalIgnoreCase));
                relationship.SetAttributeValue("Target", "/" + targetName);
            });
            RewriteXmlEntry(archive, "[Content_Types].xml", document => {
                XNamespace contentTypes = "http://schemas.openxmlformats.org/package/2006/content-types";
                XElement contentType = document.Descendants(contentTypes + "Override")
                    .Single(element => string.Equals((string?)element.Attribute("PartName"), generatedUri, StringComparison.OrdinalIgnoreCase));
                contentType.SetAttributeValue("PartName", "/" + targetName);
            });
            return "/" + targetName;
        }

        private static void RewriteXmlEntry(ZipArchive archive, string entryName, Action<XDocument> update) {
            ZipArchiveEntry entry = archive.GetEntry(entryName)
                ?? throw new InvalidDataException("Required OPC metadata entry was not found: " + entryName + ".");
            XDocument document;
            using (Stream input = entry.Open()) {
                using XmlReader reader = XmlReader.Create(input, SafeXmlReaderSettings());
                document = XDocument.Load(reader, LoadOptions.PreserveWhitespace);
            }
            entry.Delete();
            ZipArchiveEntry replacement = archive.CreateEntry(entryName, CompressionLevel.Optimal);
            using Stream output = replacement.Open();
            using XmlWriter writer = XmlWriter.Create(output, new XmlWriterSettings {
                Encoding = new System.Text.UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
                Indent = false,
                OmitXmlDeclaration = false
            });
            update(document);
            document.Save(writer);
        }

        private static List<string> ResolvePartUris(
            OfficePackageSignatureArchive package,
            OfficePackageSigningOptions options,
            out IReadOnlyList<string> missingPartUris) {
            HashSet<string>? requested = options.PartUris == null
                ? null
                : new HashSet<string>(options.PartUris.Select(OfficePackageSignatureArchive.NormalizePartUri), StringComparer.OrdinalIgnoreCase);
            var result = package.PartUris
                .Where(uri => !IsSignaturePart(uri) && !IsRelationshipPart(uri))
                .Where(uri => requested == null || requested.Contains(uri))
                .OrderBy(uri => uri, StringComparer.OrdinalIgnoreCase)
                .ToList();
            missingPartUris = requested == null
                ? Array.Empty<string>()
                : requested.Where(uri => !result.Contains(uri, StringComparer.OrdinalIgnoreCase))
                    .OrderBy(uri => uri, StringComparer.OrdinalIgnoreCase)
                    .ToArray();
            return result;
        }

        private static XElement? CreateRelationshipReference(
            OfficePackageSignatureArchive package,
            string relationshipPartUri,
            OfficePackageSigningOptions options,
            ref int remainingRelationshipSelectors,
            ref long totalDigestBytes,
            out int selectorCount) {
            selectorCount = 0;
            if (!package.ContainsPart(relationshipPartUri)) return null;
            ReserveDigestBytes(package, relationshipPartUri, options, ref totalDigestBytes);
            byte[] bytes = package.ReadPart(relationshipPartUri, options.MaxPartBytes);
            var ids = new List<string>();
            using (var stream = new MemoryStream(bytes, writable: false)) {
                using XmlReader reader = XmlReader.Create(stream, SafeXmlReaderSettings());
                const string relationshipNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";
                while (reader.Read()) {
                    if (reader.NodeType != XmlNodeType.Element ||
                        reader.LocalName != "Relationship" ||
                        reader.NamespaceURI != relationshipNamespace) continue;
                    if (IsSignatureRelationship(reader.GetAttribute("Type"))) continue;
                    string? id = reader.GetAttribute("Id")?.Trim();
                    if (string.IsNullOrWhiteSpace(id)) continue;
                    if (remainingRelationshipSelectors <= 0) {
                        throw new InvalidDataException(
                            "The package signature would contain more than " + options.MaxSignedReferences +
                            " authenticated references and relationship selectors.");
                    }
                    remainingRelationshipSelectors--;
                    selectorCount++;
                    ids.Add(id!);
                }
            }
            if (ids.Count == 0) return null;
            if (remainingRelationshipSelectors <= 0) {
                throw new InvalidDataException(
                    "The package signature would contain more than " + options.MaxSignedReferences +
                    " authenticated references and relationship selectors.");
            }
            remainingRelationshipSelectors--;
            ids.Sort(StringComparer.Ordinal);

            XNamespace ds = SignedXml.XmlDsigNamespaceUrl;
            XNamespace opc = DigitalSignatureNamespace;
            var relationshipTransform = new XElement(
                ds + "Transform",
                new XAttribute("Algorithm", OfficePackageSignatureArchive.RelationshipTransformAlgorithm),
                ids.Select(id => new XElement(opc + "RelationshipReference", new XAttribute("SourceId", id))));
            return CreateReference(
                OfficePackageSignatureArchive.NormalizePartUri(relationshipPartUri) + "?ContentType=" + Uri.EscapeDataString(OfficePackageSignatureArchive.RelationshipsContentType),
                options.HashAlgorithm,
                relationshipTransform,
                new XElement(ds + "Transform", new XAttribute("Algorithm", SignedXml.XmlDsigC14NTransformUrl)));
        }

        private static XElement CreateReference(
            string uri,
            string digestMethod,
            params XElement[] transforms) {
            XNamespace ds = SignedXml.XmlDsigNamespaceUrl;
            return new XElement(
                ds + "Reference",
                new XAttribute("URI", uri),
                transforms.Length == 0 ? null : new XElement(ds + "Transforms", transforms),
                new XElement(ds + "DigestMethod", new XAttribute("Algorithm", digestMethod)));
        }

        private static void ReserveDigestBytes(
            OfficePackageSignatureArchive package,
            string partUri,
            OfficePackageSigningOptions options,
            ref long totalDigestBytes) {
            if (!package.TryGetPartLength(partUri, out long partLength)) {
                throw new FileNotFoundException("The package part selected for signing was not found.", partUri);
            }
            if (partLength > options.MaxTotalDigestBytes - totalDigestBytes) {
                throw new InvalidDataException("The package parts selected for signing exceed the " + options.MaxTotalDigestBytes + " byte aggregate digest limit.");
            }
            totalDigestBytes += partLength;
        }

        private static string GetRelationshipPartUri(string partUri) {
            string normalized = OfficePackageSignatureArchive.NormalizePartUri(partUri);
            int slash = normalized.LastIndexOf('/');
            string directory = slash <= 0 ? string.Empty : normalized.Substring(0, slash);
            string name = normalized.Substring(slash + 1);
            return directory + "/_rels/" + name + ".rels";
        }

        private static bool IsRelationshipPart(string uri) =>
            uri.EndsWith(".rels", StringComparison.OrdinalIgnoreCase) &&
            (uri.Contains("/_rels/", StringComparison.OrdinalIgnoreCase) || uri.Equals("/_rels/.rels", StringComparison.OrdinalIgnoreCase));

        private static bool IsSignaturePart(string uri) =>
            uri.StartsWith("/_xmlsignatures/", StringComparison.OrdinalIgnoreCase) ||
            uri.StartsWith("/package/services/digital-signature/", StringComparison.OrdinalIgnoreCase);

        private static bool IsSignatureRelationship(string? relationshipType) =>
            !string.IsNullOrWhiteSpace(relationshipType) &&
            relationshipType!.IndexOf("/digital-signature/", StringComparison.OrdinalIgnoreCase) >= 0;

        private static string ResolveSignatureId(string? signatureId) {
            string resolved = string.IsNullOrWhiteSpace(signatureId)
                ? "OfficeIMOSignature" + Guid.NewGuid().ToString("N")
                : signatureId!.Trim();
            XmlConvert.VerifyNCName(resolved);
            if (string.Equals(resolved, PackageObjectId, StringComparison.Ordinal) ||
                string.Equals(resolved, SignatureTimePropertyId, StringComparison.Ordinal)) {
                throw new ArgumentException("SignatureId is reserved for an internal OPC signature node.", nameof(signatureId));
            }
            return resolved;
        }

        private static string ResolveSignatureMethod(string digestMethod) {
            switch (digestMethod.Trim()) {
                case SignedXml.XmlDsigSHA1Url:
                    return SignedXml.XmlDsigRSASHA1Url;
                case "http://www.w3.org/2001/04/xmlenc#sha256":
                case "http://www.w3.org/2001/04/xmldsig-more#sha256":
                    return SignedXml.XmlDsigRSASHA256Url;
                case "http://www.w3.org/2001/04/xmldsig-more#sha384":
                    return SignedXml.XmlDsigRSASHA384Url;
                case "http://www.w3.org/2001/04/xmlenc#sha512":
                case "http://www.w3.org/2001/04/xmldsig-more#sha512":
                    return SignedXml.XmlDsigRSASHA512Url;
                default:
                    throw new NotSupportedException("The OPC package signature hash algorithm is not supported: " + digestMethod + ".");
            }
        }

        private static XmlDocument CreateXmlDocument() => new() {
            PreserveWhitespace = true,
            XmlResolver = null
        };

        private static XmlReaderSettings SafeXmlReaderSettings() => new() {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            MaxCharactersInDocument = 64L * 1024 * 1024
        };

        private static OfficePackageSigningResult Failed(string filePath, string detail) =>
            new(
                filePath,
                isSupported: true,
                succeeded: false,
                signedPartCount: 0,
                signedRelationshipSelectorCount: 0,
                signatureCount: 0,
                signaturePartUri: null,
                details: new[] { detail });

        private static bool IsSigningException(Exception exception) =>
            exception is IOException or UnauthorizedAccessException or InvalidOperationException or ArgumentException or
                InvalidDataException or CryptographicException or XmlException or NotSupportedException;

        private sealed class SignatureBoundedMemoryStream : MemoryStream {
            private readonly long _maxLength;

            internal SignatureBoundedMemoryStream(long maxLength) {
                _maxLength = maxLength;
            }

            public override void Write(byte[] buffer, int offset, int count) {
                EnsureWithinLimit(checked(Position + count));
                base.Write(buffer, offset, count);
            }

            public override void WriteByte(byte value) {
                EnsureWithinLimit(checked(Position + 1));
                base.WriteByte(value);
            }

            public override void SetLength(long value) {
                EnsureWithinLimit(value);
                base.SetLength(value);
            }

            private void EnsureWithinLimit(long length) {
                if (length > _maxLength) {
                    throw new InvalidDataException("The generated signature XML exceeds the " + _maxLength + " byte signing limit.");
                }
            }
        }

        private readonly struct SigningPayload {
            internal SigningPayload(byte[] signatureXml, int signedPartCount, int relationshipSelectorCount) {
                SignatureXml = signatureXml;
                SignedPartCount = signedPartCount;
                RelationshipSelectorCount = relationshipSelectorCount;
            }

            internal byte[] SignatureXml { get; }
            internal int SignedPartCount { get; }
            internal int RelationshipSelectorCount { get; }
        }

        private readonly struct SignaturePartWriteResult {
            internal SignaturePartWriteResult(string signaturePartUri, int signatureCount) {
                SignaturePartUri = signaturePartUri;
                SignatureCount = signatureCount;
            }

            internal string SignaturePartUri { get; }
            internal int SignatureCount { get; }
        }
    }
}
