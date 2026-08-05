#nullable enable
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing.Internal;
using System.Diagnostics.CodeAnalysis;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Xml;
using System.Xml.Linq;
using OfficeIMO.Security;

namespace OfficeIMO.Security {
    /// <summary>Options for signing an Open Packaging Convention package.</summary>
    public sealed class OfficePackageSigningOptions {
        internal const string Sha256HashAlgorithm = "http://www.w3.org/2001/04/xmlenc#sha256";

        /// <summary>Package-part URIs to sign, or <see langword="null"/> to sign every eligible package part.</summary>
        public IReadOnlyCollection<string>? PartUris { get; set; }
        /// <summary>Whether to authenticate selected root-package relationships.</summary>
        public bool IncludePackageRelationships { get; set; } = true;
        /// <summary>Whether to authenticate relationships owned by signed parts.</summary>
        public bool IncludePartRelationships { get; set; } = true;
        /// <summary>XML DSig digest algorithm URI. The implementation accepts its bounded supported set only.</summary>
        public string HashAlgorithm { get; set; } = Sha256HashAlgorithm;
        /// <summary>Optional XML signature identifier.</summary>
        public string? SignatureId { get; set; }
        /// <summary>Optional signing time; UTC now is used when omitted.</summary>
        public DateTimeOffset? SigningTime { get; set; }
        /// <summary>Additional certificates to embed after the signing certificate.</summary>
        public IReadOnlyCollection<X509Certificate2>? AdditionalCertificates { get; set; }
        /// <summary>Maximum number of package parts accepted for signing.</summary>
        public int MaxPackageParts { get; set; } = 10000;
        /// <summary>Maximum package size in bytes accepted for signing.</summary>
        public long MaxPackageBytes { get; set; } = 512L * 1024 * 1024;
        /// <summary>Maximum uncompressed size in bytes of one signed part.</summary>
        public long MaxPartBytes { get; set; } = 256L * 1024 * 1024;
        /// <summary>Maximum aggregate bytes read while calculating package digests.</summary>
        public long MaxTotalDigestBytes { get; set; } = 512L * 1024 * 1024;
        /// <summary>Maximum authenticated part references and relationship selectors.</summary>
        public int MaxSignedReferences { get; set; } = 4096;
        /// <summary>Maximum generated XML signature size in bytes.</summary>
        public long MaxSignatureBytes { get; set; } = 16L * 1024 * 1024;
        /// <summary>Maximum number of certificates embedded in the signature.</summary>
        public int MaxCertificates { get; set; } = 64;
        /// <summary>Maximum encoded size in bytes of one certificate.</summary>
        public long MaxCertificateBytes { get; set; } = 4L * 1024 * 1024;
        /// <summary>Maximum aggregate encoded certificate bytes.</summary>
        public long MaxTotalCertificateBytes { get; set; } = 64L * 1024 * 1024;
        internal Action<string>? BeforeValidation { get; set; }
        internal Func<string, string, int, string?>? ValidateBeforeCommit { get; set; }
        internal Action<string, string>? BeforeCommit { get; set; }
    }

    /// <summary>Result of an attempted Open Packaging Convention package-signing operation.</summary>
    public sealed class OfficePackageSigningResult {
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

        /// <summary>Absolute path of the requested package.</summary>
        public string FilePath { get; }
        /// <summary>Whether the current target framework supports package signing.</summary>
        public bool IsSupported { get; }
        /// <summary>Whether the signature was created, read back, and atomically committed.</summary>
        public bool Succeeded { get; }
        /// <summary>Number of package parts selected for signing.</summary>
        public int SignedPartCount { get; }
        /// <summary>Number of relationship selectors included in the signature manifest.</summary>
        public int SignedRelationshipSelectorCount { get; }
        /// <summary>Number of signature relationships attached to the reused signature origin after signing.</summary>
        public int SignatureCount { get; }
        /// <summary>URI of the newly created signature part, when successful.</summary>
        public string? SignaturePartUri { get; }
        /// <summary>Deterministic success evidence or failure details.</summary>
        public IReadOnlyList<string> Details { get; }
    }

    /// <summary>Creates interoperable OPC XML signatures through cross-platform cryptographic primitives.</summary>
    internal static class OfficePackageSignatureWriter {
        private const string DigitalSignatureNamespace = "http://schemas.openxmlformats.org/package/2006/digital-signature";
        private const string SignatureOriginRelationship = "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin";
        private const string SignatureRelationship = "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature";
        private const string SignatureCertificateRelationship = "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/certificate";
        private const string SignatureContentType = "application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml";
        private const string ObjectReferenceType = "http://www.w3.org/2000/09/xmldsig#Object";
        private const string PackageObjectId = "idPackageObject";
        private const string SignatureTimePropertyId = "idSignatureTime";

        internal static OfficePackageSigningResult Sign(
            string filePath,
            X509Certificate2 certificate,
            IOfficeSecurityProvider securityProvider,
            OfficePackageSigningOptions? options = null) {
            options ??= new OfficePackageSigningOptions();
            if (string.IsNullOrWhiteSpace(filePath)) return Failed(filePath ?? string.Empty, "A package path is required.");

            string fullPath = Path.GetFullPath(filePath);
            if (!File.Exists(fullPath)) return Failed(fullPath, "The package file does not exist.");
            if (certificate == null) return Failed(fullPath, "A signing certificate is required.");
            if (securityProvider == null) return Failed(fullPath, "An OfficeIMO security provider is required.");
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
                IReadOnlyList<X509Certificate2> signingCertificates = ValidateSigningCertificates(certificate, options);

                stagingPath = OfficeFileCommit.CreateStagingPath(fullPath);
                OfficePackageFileSnapshot.CopyBounded(fullPath, stagingPath, options.MaxPackageBytes);
                string sourcePackageHash = OfficePackageFileSnapshot.ComputeSha256(stagingPath, options.MaxPackageBytes);
                EnsurePackagePartCountWithinLimit(stagingPath, options);
                OfficePackageSignatureInfo existingSignatures = OfficePackageSignatureService.Inspect(
                    stagingPath,
                    new OfficePackageSignatureInspectionOptions { VerifyDigests = false });
                if (!existingSignatures.HasSignatures) PrepareDigitalSignatureMetadata(stagingPath);
                byte[] packageBytes = File.ReadAllBytes(stagingPath);
                SigningPayload payload = CreateSignature(packageBytes, signingCertificates, securityProvider, options);
                SignaturePartWriteResult write = AddSignaturePart(stagingPath, payload.SignatureXml);
                EnsureSignedPackageWithinLimits(stagingPath, options);
                options.BeforeValidation?.Invoke(stagingPath);
                string validationInputHash = OfficePackageFileSnapshot.ComputeSha256(stagingPath, options.MaxPackageBytes);
                string? validationFailure = options.ValidateBeforeCommit?.Invoke(
                    stagingPath,
                    write.SignaturePartUri,
                    write.SignatureCount);
                string validatedPackageHash = OfficePackageFileSnapshot.ComputeSha256(stagingPath, options.MaxPackageBytes);
                if (!string.Equals(validationInputHash, validatedPackageHash, StringComparison.Ordinal)) {
                    return Failed(fullPath, "The staging package changed during validation readback; the original source was preserved.");
                }
                if (!string.IsNullOrWhiteSpace(validationFailure)) {
                    return Failed(fullPath, validationFailure!);
                }
                options.BeforeCommit?.Invoke(stagingPath, fullPath);
                if (!string.Equals(
                    validatedPackageHash,
                    OfficePackageFileSnapshot.ComputeSha256(stagingPath, options.MaxPackageBytes),
                    StringComparison.Ordinal)) {
                    return Failed(fullPath, "The validated staging package changed before atomic commit; the original source was preserved.");
                }
                if (!OfficeFileCommit.TryCommitTemporaryFileAtomicallyIfDestinationUnchanged(
                    stagingPath,
                    fullPath,
                    displacedPath => string.Equals(
                        sourcePackageHash,
                        OfficePackageFileSnapshot.ComputeSha256(displacedPath, options.MaxPackageBytes),
                        StringComparison.Ordinal),
                    installedPath => string.Equals(
                        validatedPackageHash,
                        OfficePackageFileSnapshot.ComputeSha256(installedPath, options.MaxPackageBytes),
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
            EnsurePackagePartCountWithinLimit(packagePath, options);
        }

        private static void EnsurePackagePartCountWithinLimit(
            string packagePath,
            OfficePackageSigningOptions options) {
            using var archive = new OfficePackageSignatureArchive(
                File.ReadAllBytes(packagePath),
                options.MaxPackageParts,
                options.MaxPartBytes);
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
            using var archive = ZipFile.Open(packagePath, ZipArchiveMode.Update);
            ZipArchiveEntry? entry = archive.GetEntry("docProps/app.xml");
            if (entry == null || entry.Length == 0) return;
            XDocument document;
            try {
                document = ReadXmlEntry(entry);
            } catch (XmlException) {
                return;
            }
            XElement? root = document.Root;
            if (root == null || root.Elements().Any(element => element.Name.LocalName == "DigitalSignature")) return;
            root.Add(new XElement(root.Name.Namespace + "DigitalSignature"));
            WriteXmlEntry(archive, "docProps/app.xml", document);
        }

        [UnconditionalSuppressMessage("Trimming", "IL2026", Justification = "Package signing selects a closed RSA and canonicalization algorithm set in code; it does not resolve algorithm implementations from caller-supplied XML.")]
        [UnconditionalSuppressMessage("AOT", "IL3050", Justification = "Package signing never uses the XSLT XML DSig transform and does not compile dynamic XSLT code.")]
        private static SigningPayload CreateSignature(
            byte[] packageBytes,
            IReadOnlyList<X509Certificate2> signingCertificates,
            IOfficeSecurityProvider securityProvider,
            OfficePackageSigningOptions options) {
            using var package = new OfficePackageSignatureArchive(
                packageBytes,
                options.MaxPackageParts,
                options.MaxPartBytes,
                securityProvider);
            List<string> partUris = ResolvePartUris(package, options, out IReadOnlyList<string> missingPartUris);
            if (missingPartUris.Count > 0) {
                throw new InvalidOperationException("Requested signing part(s) were not found: " + string.Join(", ", missingPartUris) + ".");
            }
            if (partUris.Count == 0) throw new InvalidOperationException("No package parts were selected for signing.");

            string signatureId = ResolveSignatureId(options.SignatureId);
            XNamespace ds = XmlDigitalSignatureAlgorithms.Namespace;
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

            byte[] objectXml = System.Text.Encoding.UTF8.GetBytes(
                new XElement("Root", manifest, signatureProperties).ToString(SaveOptions.DisableFormatting));
            if (objectXml.LongLength > options.MaxSignatureBytes) {
                throw new InvalidDataException(
                    "The generated signature XML exceeds the " + options.MaxSignatureBytes + " byte signing limit.");
            }
            var request = new XmlDigitalSignatureCreationRequest(
                objectXml,
                signingCertificates[0],
                signatureId,
                PackageObjectId,
                ObjectReferenceType,
                XmlDigitalSignatureAlgorithms.CanonicalXml,
                ResolveSignatureMethod(options.HashAlgorithm),
                options.HashAlgorithm) {
                AdditionalCertificates = signingCertificates.Skip(1).ToArray(),
                MaxObjectBytes = options.MaxSignatureBytes,
                MaxOutputBytes = options.MaxSignatureBytes
            };
            byte[] signatureXml = securityProvider.CreateXmlSignature(request);
            return new SigningPayload(signatureXml, partUris.Count, relationshipSelectorCount);
        }

        private static SignaturePartWriteResult AddSignaturePart(string packagePath, byte[] signatureXml) {
            using var archive = ZipFile.Open(packagePath, ZipArchiveMode.Update);
            XNamespace relationships = "http://schemas.openxmlformats.org/package/2006/relationships";
            XDocument rootRelationships = ReadOrCreateRelationships(archive.GetEntry("_rels/.rels"), relationships);
            XElement root = rootRelationships.Root!;
            XElement? originRelationship = root.Elements(relationships + "Relationship").FirstOrDefault(element =>
                string.Equals((string?)element.Attribute("Type"), SignatureOriginRelationship, StringComparison.Ordinal) &&
                !string.Equals((string?)element.Attribute("TargetMode"), "External", StringComparison.OrdinalIgnoreCase));
            string? existingOriginTarget = (string?)originRelationship?.Attribute("Target");
            string originName = string.IsNullOrWhiteSpace(existingOriginTarget)
                ? "_xmlsignatures/origin.sigs"
                : OfficePackageSignatureArchive.NormalizePartUri(existingOriginTarget!).TrimStart('/');
            if (archive.GetEntry(originName) == null) {
                originName = "_xmlsignatures/origin.sigs";
                originRelationship = null;
            }
            string originRelationshipsName = RelationshipPartName(originName);
            int signatureIndex = 1;
            string signatureName;
            do {
                signatureName = "_xmlsignatures/sig" + signatureIndex.ToString(System.Globalization.CultureInfo.InvariantCulture) + ".xml";
                signatureIndex++;
            } while (archive.GetEntry(signatureName) != null);

            ZipArchiveEntry signatureEntry = archive.CreateEntry(signatureName, CompressionLevel.Optimal);
            using (Stream output = signatureEntry.Open()) output.Write(signatureXml, 0, signatureXml.Length);
            if (archive.GetEntry(originName) == null) archive.CreateEntry(originName, CompressionLevel.NoCompression);

            if (originRelationship == null) {
                root.Add(new XElement(relationships + "Relationship",
                    new XAttribute("Id", NextRelationshipId(root, relationships)),
                    new XAttribute("Type", SignatureOriginRelationship),
                    new XAttribute("Target", originName)));
            }
            WriteXmlEntry(archive, "_rels/.rels", rootRelationships);

            XDocument originRelationships = ReadOrCreateRelationships(archive.GetEntry(originRelationshipsName), relationships);
            XElement originRoot = originRelationships.Root!;
            originRoot.Add(new XElement(relationships + "Relationship",
                new XAttribute("Id", NextRelationshipId(originRoot, relationships)),
                new XAttribute("Type", SignatureRelationship),
                new XAttribute("Target", RelativePartTarget(originName, signatureName))));
            WriteXmlEntry(archive, originRelationshipsName, originRelationships);

            ZipArchiveEntry typesEntry = archive.GetEntry("[Content_Types].xml")
                ?? throw new InvalidDataException("The OPC package is missing [Content_Types].xml.");
            XDocument contentTypesDocument = ReadXmlEntry(typesEntry);
            XNamespace contentTypes = "http://schemas.openxmlformats.org/package/2006/content-types";
            XElement typesRoot = contentTypesDocument.Root
                ?? throw new InvalidDataException("The OPC content-types document has no root element.");
            EnsureOverride(typesRoot, contentTypes, "/" + originName,
                "application/vnd.openxmlformats-package.digital-signature-origin");
            EnsureOverride(typesRoot, contentTypes, "/" + signatureName,
                SignatureContentType);
            WriteXmlEntry(archive, "[Content_Types].xml", contentTypesDocument);

            int signatureCount = originRoot.Elements(relationships + "Relationship").Count(element =>
                string.Equals((string?)element.Attribute("Type"), SignatureRelationship, StringComparison.Ordinal) &&
                !string.Equals((string?)element.Attribute("TargetMode"), "External", StringComparison.OrdinalIgnoreCase));
            return new SignaturePartWriteResult("/" + signatureName, signatureCount);
        }

        private static string RelationshipPartName(string partName) {
            int slash = partName.LastIndexOf('/');
            string directory = slash < 0 ? string.Empty : partName.Substring(0, slash + 1);
            string fileName = slash < 0 ? partName : partName.Substring(slash + 1);
            return directory + "_rels/" + fileName + ".rels";
        }

        private static string RelativePartTarget(string sourcePartName, string targetPartName) {
            var source = new Uri("http://officeimo.package/" + sourcePartName, UriKind.Absolute);
            var target = new Uri("http://officeimo.package/" + targetPartName, UriKind.Absolute);
            return Uri.UnescapeDataString(source.MakeRelativeUri(target).ToString());
        }

        private static void EnsureOverride(XElement root, XNamespace contentTypes, string partName, string contentType) {
            XElement? existing = root.Elements(contentTypes + "Override").FirstOrDefault(element =>
                string.Equals((string?)element.Attribute("PartName"), partName, StringComparison.OrdinalIgnoreCase));
            if (existing == null) {
                root.Add(new XElement(contentTypes + "Override",
                    new XAttribute("PartName", partName), new XAttribute("ContentType", contentType)));
            } else {
                existing.SetAttributeValue("ContentType", contentType);
            }
        }

        private static XDocument ReadOrCreateRelationships(ZipArchiveEntry? entry, XNamespace relationships) =>
            entry == null
                ? new XDocument(new XElement(relationships + "Relationships"))
                : ReadXmlEntry(entry);

        private static string NextRelationshipId(XElement root, XNamespace relationships) {
            var ids = new HashSet<string>(root.Elements(relationships + "Relationship")
                .Select(element => (string?)element.Attribute("Id"))
                .Where(id => !string.IsNullOrWhiteSpace(id))
                .Select(id => id!), StringComparer.Ordinal);
            for (int index = 1; ; index++) {
                string candidate = "rId" + index.ToString(System.Globalization.CultureInfo.InvariantCulture);
                if (!ids.Contains(candidate)) return candidate;
            }
        }

        private static XDocument ReadXmlEntry(ZipArchiveEntry entry) {
            using Stream input = entry.Open();
            using XmlReader reader = XmlReader.Create(input, SafeXmlReaderSettings());
            return XDocument.Load(reader, LoadOptions.PreserveWhitespace);
        }

        private static void WriteXmlEntry(ZipArchive archive, string entryName, XDocument document) {
            archive.GetEntry(entryName)?.Delete();
            ZipArchiveEntry replacement = archive.CreateEntry(entryName, CompressionLevel.Optimal);
            using Stream output = replacement.Open();
            using XmlWriter writer = XmlWriter.Create(output, new XmlWriterSettings {
                Encoding = new System.Text.UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
                Indent = false,
                OmitXmlDeclaration = false
            });
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

            XNamespace ds = XmlDigitalSignatureAlgorithms.Namespace;
            XNamespace opc = DigitalSignatureNamespace;
            var relationshipTransform = new XElement(
                ds + "Transform",
                new XAttribute("Algorithm", OfficePackageSignatureArchive.RelationshipTransformAlgorithm),
                ids.Select(id => new XElement(opc + "RelationshipReference", new XAttribute("SourceId", id))));
            return CreateReference(
                OfficePackageSignatureArchive.NormalizePartUri(relationshipPartUri) + "?ContentType=" + Uri.EscapeDataString(OfficePackageSignatureArchive.RelationshipsContentType),
                options.HashAlgorithm,
                relationshipTransform,
                new XElement(ds + "Transform", new XAttribute("Algorithm", XmlDigitalSignatureAlgorithms.CanonicalXml)));
        }

        private static XElement CreateReference(
            string uri,
            string digestMethod,
            params XElement[] transforms) {
            XNamespace ds = XmlDigitalSignatureAlgorithms.Namespace;
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
            (uri.IndexOf("/_rels/", StringComparison.OrdinalIgnoreCase) >= 0 || uri.Equals("/_rels/.rels", StringComparison.OrdinalIgnoreCase));

        private static bool IsSignaturePart(string uri) =>
            uri.StartsWith("/_xmlsignatures/", StringComparison.OrdinalIgnoreCase) ||
            uri.StartsWith("/package/services/digital-signature/", StringComparison.OrdinalIgnoreCase);

        private static bool IsSignatureRelationship(string? relationshipType) =>
            string.Equals(relationshipType, SignatureOriginRelationship, StringComparison.Ordinal) ||
            string.Equals(relationshipType, SignatureRelationship, StringComparison.Ordinal) ||
            string.Equals(relationshipType, SignatureCertificateRelationship, StringComparison.Ordinal);

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
                case "http://www.w3.org/2000/09/xmldsig#sha1":
                    return XmlDigitalSignatureAlgorithms.RsaSha1;
                case "http://www.w3.org/2001/04/xmlenc#sha256":
                case "http://www.w3.org/2001/04/xmldsig-more#sha256":
                    return XmlDigitalSignatureAlgorithms.RsaSha256;
                case "http://www.w3.org/2001/04/xmldsig-more#sha384":
                    return XmlDigitalSignatureAlgorithms.RsaSha384;
                case "http://www.w3.org/2001/04/xmlenc#sha512":
                case "http://www.w3.org/2001/04/xmldsig-more#sha512":
                    return XmlDigitalSignatureAlgorithms.RsaSha512;
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
