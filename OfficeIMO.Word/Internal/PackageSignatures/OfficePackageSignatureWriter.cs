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
            long packageLength = new FileInfo(fullPath).Length;
            if (packageLength > options.MaxPackageBytes) {
                return Failed(fullPath, "The package exceeds the " + options.MaxPackageBytes + " byte signing limit.");
            }

            string stagingPath = string.Empty;
            try {
                if (!certificate.HasPrivateKey) return Failed(fullPath, "The signing certificate must include a private key.");
                using RSA? signingKey = certificate.GetRSAPrivateKey();
                if (signingKey == null) return Failed(fullPath, "OPC package signing requires an RSA certificate with an accessible private key.");

                stagingPath = OfficeFileCommit.CreateStagingPath(fullPath);
                File.Copy(fullPath, stagingPath, overwrite: false);
                PrepareDigitalSignatureMetadata(stagingPath);
                byte[] packageBytes = File.ReadAllBytes(stagingPath);
                SigningPayload payload = CreateSignature(packageBytes, certificate, signingKey, options);
                SignaturePartWriteResult write = AddSignaturePart(stagingPath, payload.SignatureXml);
                EnsureSignedPackageWithinLimits(stagingPath, options);
                OfficeFileCommit.CommitTemporaryFileAtomically(stagingPath, fullPath);
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
            X509Certificate2 certificate,
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
            foreach (string partUri in partUris) {
                if (!package.TryGetContentType(partUri, out string contentType)) {
                    throw new InvalidDataException("The OPC package does not declare a content type for " + partUri + ".");
                }
                manifest.Add(CreateReference(partUri + "?ContentType=" + Uri.EscapeDataString(contentType), options.HashAlgorithm));
            }

            int relationshipSelectorCount = 0;
            if (options.IncludePackageRelationships) {
                XElement? packageRelationships = CreateRelationshipReference(package, "/_rels/.rels", options);
                if (packageRelationships != null) {
                    relationshipSelectorCount += CountRelationshipSelectors(packageRelationships);
                    manifest.Add(packageRelationships);
                }
            }
            if (options.IncludePartRelationships) {
                foreach (string partUri in partUris) {
                    string relationshipPartUri = GetRelationshipPartUri(partUri);
                    XElement? partRelationships = CreateRelationshipReference(package, relationshipPartUri, options);
                    if (partRelationships == null) continue;
                    relationshipSelectorCount += CountRelationshipSelectors(partRelationships);
                    manifest.Add(partRelationships);
                }
            }

            long totalDigestBytes = 0;
            foreach (XElement reference in manifest.Elements(ds + "Reference")) {
                string? referenceUri = ((string?)reference.Attribute("URI"))?.Trim();
                string targetPartUri = OfficePackageSignatureArchive.NormalizeReferencePartUri(referenceUri)
                    ?? throw new InvalidDataException("The OPC signature Reference is not a package-part URI: " + referenceUri + ".");
                if (!package.TryGetPartLength(targetPartUri, out long partLength)) {
                    throw new FileNotFoundException("The package part selected for signing was not found.", targetPartUri);
                }
                if (partLength > options.MaxTotalDigestBytes - totalDigestBytes) {
                    throw new InvalidDataException("The package parts selected for signing exceed the " + options.MaxTotalDigestBytes + " byte aggregate digest limit.");
                }
                totalDigestBytes += partLength;
                reference.Add(new XElement(ds + "DigestValue", package.ComputeDigestValue(reference, options.MaxPartBytes)));
            }

            DateTimeOffset signingTime = options.SigningTime ?? DateTimeOffset.UtcNow;
            var signatureProperties = new XElement(
                ds + "SignatureProperties",
                new XElement(
                    ds + "SignatureProperty",
                    new XAttribute("Id", "idSignatureTime"),
                    new XAttribute("Target", "#" + signatureId),
                    new XElement(
                        opc + "SignatureTime",
                        new XElement(opc + "Format", "YYYY-MM-DDThh:mm:ss.sTZD"),
                        new XElement(opc + "Value", signingTime.ToString("yyyy-MM-dd'T'HH:mm:ss.FFFFFFFK", System.Globalization.CultureInfo.InvariantCulture)))));

            XmlDocument objectDocument = CreateXmlDocument();
            objectDocument.LoadXml(new XElement("Root", manifest, signatureProperties).ToString(SaveOptions.DisableFormatting));
            var packageObject = new DataObject {
                Id = "idPackageObject",
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
            signedXml.AddReference(new Reference("#idPackageObject") {
                Type = ObjectReferenceType,
                DigestMethod = options.HashAlgorithm
            });

            var keyInfo = new KeyInfo();
            var x509Data = new KeyInfoX509Data(certificate);
            if (options.AdditionalCertificates != null) {
                foreach (X509Certificate2 additional in options.AdditionalCertificates) {
                    if (additional == null || string.Equals(additional.Thumbprint, certificate.Thumbprint, StringComparison.OrdinalIgnoreCase)) continue;
                    x509Data.AddCertificate(additional);
                }
            }
            keyInfo.AddClause(x509Data);
            signedXml.KeyInfo = keyInfo;
            signedXml.ComputeSignature();

            XmlElement signature = signedXml.GetXml();
            signatureDocument.AppendChild(signatureDocument.ImportNode(signature, deep: true));
            using var output = new MemoryStream();
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
            OfficePackageSigningOptions options) {
            if (!package.ContainsPart(relationshipPartUri)) return null;
            byte[] bytes = package.ReadPart(relationshipPartUri, options.MaxPartBytes);
            XDocument relationships;
            using (var stream = new MemoryStream(bytes, writable: false)) {
                using XmlReader reader = XmlReader.Create(stream, SafeXmlReaderSettings());
                relationships = XDocument.Load(reader, LoadOptions.PreserveWhitespace);
            }
            XNamespace rel = "http://schemas.openxmlformats.org/package/2006/relationships";
            string[] ids = relationships.Root?
                .Elements(rel + "Relationship")
                .Where(element => !IsSignatureRelationship((string?)element.Attribute("Type")))
                .Select(element => ((string?)element.Attribute("Id"))?.Trim())
                .Where(id => !string.IsNullOrWhiteSpace(id))
                .Select(id => id!)
                .OrderBy(id => id, StringComparer.Ordinal)
                .ToArray() ?? Array.Empty<string>();
            if (ids.Length == 0) return null;

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

        private static int CountRelationshipSelectors(XElement reference) {
            XNamespace opc = DigitalSignatureNamespace;
            return reference.Descendants(opc + "RelationshipReference").Count() +
                   reference.Descendants(opc + "RelationshipsGroupReference").Count();
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
