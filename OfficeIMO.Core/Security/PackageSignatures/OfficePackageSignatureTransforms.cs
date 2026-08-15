#nullable enable
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using OfficeIMO.Security;

namespace OfficeIMO.Security {
    /// <summary>Reports that bounded OPC archive inspection stopped at a caller-owned resource limit.</summary>
    internal sealed class OfficePackageSignatureResourceLimitException : IOException {
        internal OfficePackageSignatureResourceLimitException(string message, long consumedBytes = 0) : base(message) {
            ConsumedBytes = consumedBytes;
        }

        internal long ConsumedBytes { get; }
    }

    /// <summary>Bounded OPC archive reader and transform-aware signature digest engine.</summary>
    internal sealed class OfficePackageSignatureArchive : IDisposable {
        internal const string RelationshipTransformAlgorithm = "http://schemas.openxmlformats.org/package/2006/RelationshipTransform";
        internal const string CanonicalXmlAlgorithm = XmlDigitalSignatureAlgorithms.CanonicalXml;
        internal const string CanonicalXmlWithCommentsAlgorithm = XmlDigitalSignatureAlgorithms.CanonicalXmlWithComments;
        internal const string RelationshipsContentType = "application/vnd.openxmlformats-package.relationships+xml";

        private readonly MemoryStream _stream;
        private readonly ZipArchive _archive;
        private readonly Dictionary<string, ZipArchiveEntry> _entries;
        private readonly Dictionary<string, string> _contentTypes;
        private readonly IOfficeSecurityProvider? _securityProvider;

        internal OfficePackageSignatureArchive(
            byte[] packageBytes,
            int maxParts = 10000,
            long maxPartBytes = 64L * 1024 * 1024,
            IOfficeSecurityProvider? securityProvider = null) {
            if (packageBytes == null) throw new ArgumentNullException(nameof(packageBytes));
            if (maxParts <= 0) throw new ArgumentOutOfRangeException(nameof(maxParts));
            if (maxPartBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maxPartBytes));
            _securityProvider = securityProvider;

            _stream = new MemoryStream(packageBytes, writable: false);
            _archive = new ZipArchive(_stream, ZipArchiveMode.Read, leaveOpen: false);
            if (_archive.Entries.Count > maxParts) {
                throw new OfficePackageSignatureResourceLimitException(
                    "The OPC package contains more than " + maxParts + " ZIP entries.");
            }

            _entries = new Dictionary<string, ZipArchiveEntry>(StringComparer.OrdinalIgnoreCase);
            foreach (ZipArchiveEntry entry in _archive.Entries) {
                if (entry.Name.Length == 0) continue;
                string uri = NormalizePartUri(entry.FullName);
                if (_entries.ContainsKey(uri)) {
                    throw new InvalidDataException("The OPC package contains duplicate part URI '" + uri + "'.");
                }
                _entries.Add(uri, entry);
            }

            _contentTypes = ReadContentTypes(maxPartBytes);
        }

        internal IReadOnlyList<string> PartUris => _entries.Keys
            .Where(uri => !uri.Equals("/[Content_Types].xml", StringComparison.OrdinalIgnoreCase))
            .OrderBy(uri => uri, StringComparer.OrdinalIgnoreCase)
            .ToArray();

        internal bool ContainsPart(string partUri) => _entries.ContainsKey(NormalizePartUri(partUri));

        internal bool TryGetPartLength(string partUri, out long length) {
            if (_entries.TryGetValue(NormalizePartUri(partUri), out ZipArchiveEntry? entry)) {
                length = entry.Length;
                return true;
            }
            length = 0;
            return false;
        }

        internal bool TryGetContentType(string partUri, out string contentType) =>
            _contentTypes.TryGetValue(NormalizePartUri(partUri), out contentType!);

        internal byte[] ReadPart(string partUri, long maxBytes) {
            string normalized = NormalizePartUri(partUri);
            if (!_entries.TryGetValue(normalized, out ZipArchiveEntry? entry)) {
                throw new FileNotFoundException("The OPC package part was not found.", normalized);
            }
            if (maxBytes < 0) throw new ArgumentOutOfRangeException(nameof(maxBytes));
            if (entry.Length > maxBytes) {
                throw new InvalidDataException("The OPC package part '" + normalized + "' exceeds the " + maxBytes + " byte limit.");
            }

            using Stream source = entry.Open();
            using var output = new MemoryStream(entry.Length > int.MaxValue ? 0 : (int)entry.Length);
            CopyBounded(source, output, maxBytes, normalized);
            return output.ToArray();
        }

        internal OfficePackageDigestResult VerifyReference(
            XElement reference,
            long maxPartBytes,
            long maxDigestBytes = long.MaxValue) {
            if (maxDigestBytes < 0) {
                throw new OfficePackageSignatureResourceLimitException(
                    "OPC signature inspection exceeds the configured aggregate digest-byte limit.");
            }
            XNamespace ds = XmlDigitalSignatureAlgorithms.Namespace;
            string? uri = ((string?)reference.Attribute("URI"))?.Trim();
            string? targetPartUri = NormalizeReferencePartUri(uri);
            if (targetPartUri == null) {
                return OfficePackageDigestResult.NotChecked("Reference '" + (uri ?? string.Empty) + "' is not an OPC package-part reference.");
            }
            if (!ContainsPart(targetPartUri)) {
                return OfficePackageDigestResult.Failed("Digest verification failed because package part " + targetPartUri + " does not exist.");
            }
            if (TryGetPartLength(targetPartUri, out long targetPartLength) && targetPartLength > maxPartBytes) {
                return OfficePackageDigestResult.Unsupported(
                    "Digest verification for " + targetPartUri +
                    " was not checked because the package part exceeds the " + maxPartBytes + " byte limit.");
            }
            string? declaredContentType;
            try {
                declaredContentType = GetDeclaredContentType(uri);
            } catch (UriFormatException exception) {
                return OfficePackageDigestResult.Failed("Digest verification failed because Reference URI content-type encoding is invalid: " + exception.Message);
            }
            if (string.IsNullOrWhiteSpace(declaredContentType)) {
                return OfficePackageDigestResult.Failed(
                    "Digest verification failed because package part " + targetPartUri +
                    " is not bound to an OPC content type in its Reference URI.");
            }
            if (!TryGetContentType(targetPartUri, out string actualContentType)) {
                return OfficePackageDigestResult.Failed("Digest verification failed because package part " + targetPartUri + " has no resolved OPC content type.");
            }
            if (!string.Equals(declaredContentType, actualContentType, StringComparison.OrdinalIgnoreCase)) {
                return OfficePackageDigestResult.Failed(
                    "Digest verification failed because package part " + targetPartUri +
                    " declares signed content type '" + declaredContentType +
                    "' but the package resolves it as '" + actualContentType + "'.");
            }

            string? digestMethod = (string?)reference.Element(ds + "DigestMethod")?.Attribute("Algorithm");
            string? digestValue = reference.Element(ds + "DigestValue")?.Value.Trim();
            if (string.IsNullOrWhiteSpace(digestMethod) || string.IsNullOrWhiteSpace(digestValue)) {
                return OfficePackageDigestResult.NotChecked("Digest verification for " + targetPartUri + " was not performed because digest metadata is incomplete.");
            }

            Func<HashAlgorithm>? hashFactory = CreateHashAlgorithm(digestMethod!);
            if (hashFactory == null) {
                return OfficePackageDigestResult.Unsupported("Digest verification for " + targetPartUri + " does not support digest method " + digestMethod + ".");
            }

            byte[] expected;
            try {
                expected = Convert.FromBase64String(digestValue!);
            } catch (FormatException exception) {
                return OfficePackageDigestResult.Failed("Digest verification for " + targetPartUri + " failed because DigestValue is not valid base64: " + exception.Message);
            }

            byte[] input;
            long transformInputBytes = 0;
            try {
                input = ApplyTransforms(targetPartUri, reference, maxPartBytes, maxDigestBytes, ref transformInputBytes);
            } catch (OfficePackageSignatureResourceLimitException exception) {
                throw new OfficePackageSignatureResourceLimitException(
                    exception.Message,
                    Math.Max(exception.ConsumedBytes, transformInputBytes));
            } catch (NotSupportedException exception) {
                return OfficePackageDigestResult.Unsupported(exception.Message, transformInputBytes);
            } catch (Exception exception) when (exception is IOException or InvalidDataException or XmlException or CryptographicException) {
                return OfficePackageDigestResult.Failed(
                    "Digest verification for " + targetPartUri + " failed while applying transforms: " + exception.Message,
                    transformInputBytes);
            }

            long digestWorkBytes;
            try {
                digestWorkBytes = AddDigestWorkBytes(transformInputBytes, input.LongLength, maxDigestBytes);
            } catch (OfficePackageSignatureResourceLimitException exception) {
                throw new OfficePackageSignatureResourceLimitException(
                    exception.Message,
                    Math.Max(exception.ConsumedBytes, transformInputBytes));
            }

            byte[] actual;
            using (HashAlgorithm hash = hashFactory()) {
                actual = hash.ComputeHash(input);
            }

            return FixedTimeEquals(actual, expected)
                ? OfficePackageDigestResult.Passed("Digest verification passed for " + targetPartUri + ".", digestWorkBytes)
                : OfficePackageDigestResult.Failed("Digest verification failed for " + targetPartUri + ".", digestWorkBytes);
        }

        internal string ComputeDigestValue(XElement reference, long maxPartBytes) {
            XNamespace ds = XmlDigitalSignatureAlgorithms.Namespace;
            string? digestMethod = (string?)reference.Element(ds + "DigestMethod")?.Attribute("Algorithm");
            if (string.IsNullOrWhiteSpace(digestMethod)) {
                throw new InvalidOperationException("The OPC signature Reference is missing DigestMethod.");
            }

            Func<HashAlgorithm>? hashFactory = CreateHashAlgorithm(digestMethod!);
            if (hashFactory == null) {
                throw new NotSupportedException("The OPC signature digest method is not supported: " + digestMethod + ".");
            }

            string? uri = ((string?)reference.Attribute("URI"))?.Trim();
            string targetPartUri = NormalizeReferencePartUri(uri)
                ?? throw new InvalidOperationException("The OPC signature Reference is not a package-part URI: " + uri + ".");
            string? declaredContentType;
            try {
                declaredContentType = GetDeclaredContentType(uri);
            } catch (UriFormatException exception) {
                throw new InvalidDataException("The OPC signature Reference content-type encoding is invalid.", exception);
            }
            if (string.IsNullOrWhiteSpace(declaredContentType)) {
                throw new InvalidDataException("The OPC signature Reference must bind the package content type for " + targetPartUri + ".");
            }
            if (!TryGetContentType(targetPartUri, out string actualContentType) ||
                !string.Equals(declaredContentType, actualContentType, StringComparison.OrdinalIgnoreCase)) {
                throw new InvalidDataException("The OPC signature Reference content type does not match the package content type for " + targetPartUri + ".");
            }
            long ignoredTransformInputBytes = 0;
            byte[] input = ApplyTransforms(targetPartUri, reference, maxPartBytes, long.MaxValue, ref ignoredTransformInputBytes);
            using HashAlgorithm hash = hashFactory();
            return Convert.ToBase64String(hash.ComputeHash(input));
        }

        public void Dispose() {
            _archive.Dispose();
            _stream.Dispose();
        }

        internal static string NormalizePartUri(string partUri) {
            if (string.IsNullOrWhiteSpace(partUri)) return "/";
            string normalized = partUri.Trim().Replace('\\', '/');
            return normalized.StartsWith("/", StringComparison.Ordinal) ? normalized : "/" + normalized;
        }

        internal static string? NormalizeReferencePartUri(string? uri) {
            if (string.IsNullOrWhiteSpace(uri)) return null;
            string trimmed = uri!.Trim();
            if (!trimmed.StartsWith("/", StringComparison.Ordinal)) return null;
            if (trimmed.IndexOf('#') >= 0) return null;
            int query = trimmed.IndexOf('?');
            if (query >= 0) trimmed = trimmed.Substring(0, query);
            return trimmed.Length == 0 ? null : NormalizePartUri(trimmed);
        }

        private static string? GetDeclaredContentType(string? uri) {
            if (string.IsNullOrWhiteSpace(uri)) return null;
            int queryStart = uri!.IndexOf('?');
            if (queryStart < 0 || queryStart == uri.Length - 1) return null;
            int fragmentStart = uri.IndexOf('#', queryStart + 1);
            string query = fragmentStart < 0
                ? uri.Substring(queryStart + 1)
                : uri.Substring(queryStart + 1, fragmentStart - queryStart - 1);
            foreach (string item in query.Split('&')) {
                int equals = item.IndexOf('=');
                string key = equals < 0 ? item : item.Substring(0, equals);
                if (!string.Equals(Uri.UnescapeDataString(key), "ContentType", StringComparison.OrdinalIgnoreCase)) continue;
                string value = equals < 0 ? string.Empty : item.Substring(equals + 1);
                return Uri.UnescapeDataString(value);
            }
            return null;
        }

        internal byte[] Canonicalize(XmlDocument document, bool includeComments = false, long maxOutputBytes = 64L * 1024L * 1024L) {
            if (_securityProvider == null) {
                throw new NotSupportedException(
                    "XML canonicalization requires an explicitly supplied OfficeIMO security provider.");
            }
            return _securityProvider.CanonicalizeXml(
                SerializeXmlForCanonicalization(document, maxOutputBytes),
                includeComments ? CanonicalXmlWithCommentsAlgorithm : CanonicalXmlAlgorithm,
                maxOutputBytes: maxOutputBytes);
        }

        private static byte[] SerializeXmlForCanonicalization(XmlDocument document, long maxBytes) {
            using var stream = new MemoryStream();
            using (XmlWriter writer = XmlWriter.Create(stream, new XmlWriterSettings {
                Encoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
                Indent = false,
                OmitXmlDeclaration = true,
                CloseOutput = false
            })) {
                foreach (XmlNode node in document.ChildNodes) {
                    if (node is not XmlDeclaration) node.WriteTo(writer);
                }
            }
            if (stream.Length > maxBytes) {
                throw new OfficePackageSignatureResourceLimitException(
                    "XML canonicalization input exceeds the " + maxBytes + " byte output-work limit.");
            }
            return stream.ToArray();
        }

        private byte[] ApplyTransforms(
            string targetPartUri,
            XElement reference,
            long maxPartBytes,
            long maxDigestBytes,
            ref long transformInputBytes) {
            XNamespace ds = XmlDigitalSignatureAlgorithms.Namespace;
            List<XElement> transforms = reference
                .Element(ds + "Transforms")?
                .Elements(ds + "Transform")
                .ToList() ?? new List<XElement>();

            if (transforms.Count == 0) {
                PreflightDigestRead(targetPartUri, maxDigestBytes, transformInputBytes);
                return ReadPart(targetPartUri, maxPartBytes);
            }

            byte[]? currentBytes = null;
            XmlDocument? currentXml = null;
            foreach (XElement transform in transforms) {
                string algorithm = ((string?)transform.Attribute("Algorithm"))?.Trim() ?? string.Empty;
                if (algorithm.Equals(RelationshipTransformAlgorithm, StringComparison.Ordinal)) {
                    if (currentBytes != null || currentXml != null) {
                        throw new NotSupportedException("RelationshipTransform must be the first transform for " + targetPartUri + ".");
                    }
                    currentXml = ApplyRelationshipTransform(
                        targetPartUri,
                        transform,
                        maxPartBytes,
                        maxDigestBytes,
                        ref transformInputBytes);
                    continue;
                }

                if (algorithm.Equals(CanonicalXmlAlgorithm, StringComparison.Ordinal) ||
                    algorithm.Equals(CanonicalXmlWithCommentsAlgorithm, StringComparison.Ordinal)) {
                    if (currentXml == null) {
                        currentXml = LoadXml(
                            currentBytes ?? ReadPartForDigestWork(
                                targetPartUri,
                                maxPartBytes,
                                maxDigestBytes,
                                ref transformInputBytes),
                            maxPartBytes);
                    }
                    currentBytes = Canonicalize(
                        currentXml,
                        algorithm.Equals(CanonicalXmlWithCommentsAlgorithm, StringComparison.Ordinal),
                        maxPartBytes);
                    currentXml = null;
                    continue;
                }

                throw new NotSupportedException("Digest verification for " + targetPartUri + " does not support transform " + algorithm + ".");
            }

            if (currentXml != null) {
                throw new NotSupportedException("RelationshipTransform for " + targetPartUri + " must be followed by an XML canonicalization transform.");
            }
            return currentBytes ?? ReadPartForDigestWork(
                targetPartUri,
                maxPartBytes,
                maxDigestBytes,
                ref transformInputBytes);
        }

        private XmlDocument ApplyRelationshipTransform(
            string targetPartUri,
            XElement transform,
            long maxPartBytes,
            long maxDigestBytes,
            ref long transformInputBytes) {
            if (!TryGetContentType(targetPartUri, out string contentType) ||
                !string.Equals(contentType, RelationshipsContentType, StringComparison.OrdinalIgnoreCase)) {
                throw new InvalidDataException(
                    "RelationshipTransform requires an OPC relationships part, but " + targetPartUri +
                    " resolves to content type '" + (contentType ?? string.Empty) + "'.");
            }
            XmlDocument source = LoadXml(
                ReadPartForDigestWork(targetPartUri, maxPartBytes, maxDigestBytes, ref transformInputBytes),
                maxPartBytes);
            const string relationshipNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";
            if (!IsValidRelationshipsDocument(source, relationshipNamespace)) {
                throw new InvalidDataException(
                    "RelationshipTransform requires a structurally valid OPC Relationships document for " + targetPartUri + ".");
            }
            XNamespace opc = "http://schemas.openxmlformats.org/package/2006/digital-signature";
            var ids = new HashSet<string>(transform
                .Elements(opc + "RelationshipReference")
                .Select(element => ((string?)element.Attribute("SourceId"))?.Trim())
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Select(value => value!), StringComparer.Ordinal);
            var types = new HashSet<string>(transform
                .Elements(opc + "RelationshipsGroupReference")
                .Select(element => ((string?)element.Attribute("SourceType"))?.Trim())
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Select(value => value!), StringComparer.Ordinal);

            XmlDocument result = CreateXmlDocument();
            XmlElement root = result.CreateElement("Relationships", relationshipNamespace);
            result.AppendChild(root);

            IEnumerable<XmlElement> relationships = source.DocumentElement?
                .ChildNodes
                .OfType<XmlElement>()
                .Where(element => element.LocalName == "Relationship" && element.NamespaceURI == relationshipNamespace)
                ?? Enumerable.Empty<XmlElement>();
            foreach (XmlElement relationship in relationships
                .Where(element => ids.Contains(element.GetAttribute("Id")) || types.Contains(element.GetAttribute("Type")))
                .OrderBy(element => element.GetAttribute("Id"), StringComparer.Ordinal)) {
                XmlElement selected = result.CreateElement("Relationship", relationshipNamespace);
                selected.SetAttribute("Id", relationship.GetAttribute("Id"));
                selected.SetAttribute("Type", relationship.GetAttribute("Type"));
                selected.SetAttribute("Target", relationship.GetAttribute("Target"));
                selected.SetAttribute("TargetMode", relationship.HasAttribute("TargetMode")
                    ? relationship.GetAttribute("TargetMode")
                    : "Internal");
                root.AppendChild(selected);
            }
            return result;
        }

        private static bool IsValidRelationshipsDocument(XmlDocument source, string relationshipNamespace) {
            XmlElement? root = source.DocumentElement;
            if (root == null || root.LocalName != "Relationships" || root.NamespaceURI != relationshipNamespace) return false;
            var ids = new HashSet<string>(StringComparer.Ordinal);
            foreach (XmlElement relationship in root.ChildNodes.OfType<XmlElement>()) {
                if (relationship.LocalName != "Relationship" || relationship.NamespaceURI != relationshipNamespace ||
                    relationship.ChildNodes.OfType<XmlElement>().Any()) return false;
                string id = relationship.GetAttribute("Id");
                string type = relationship.GetAttribute("Type");
                string target = relationship.GetAttribute("Target");
                if (string.IsNullOrWhiteSpace(id) || string.IsNullOrWhiteSpace(type) ||
                    string.IsNullOrWhiteSpace(target) || !ids.Add(id)) return false;
                if (relationship.HasAttribute("TargetMode")) {
                    string targetMode = relationship.GetAttribute("TargetMode");
                    if (targetMode != "Internal" && targetMode != "External") return false;
                }
            }
            return true;
        }

        private byte[] ReadPartForDigestWork(
            string partUri,
            long maxPartBytes,
            long maxDigestBytes,
            ref long transformInputBytes) {
            if (TryGetPartLength(partUri, out long declaredLength) &&
                declaredLength > maxDigestBytes - transformInputBytes) {
                throw new OfficePackageSignatureResourceLimitException(
                    "OPC signature inspection exceeds the configured aggregate digest-byte limit.");
            }
            byte[] bytes = ReadPart(partUri, maxPartBytes);
            transformInputBytes = AddDigestWorkBytes(transformInputBytes, bytes.LongLength, maxDigestBytes);
            return bytes;
        }

        private void PreflightDigestRead(string partUri, long maximumBytes, long currentBytes) {
            if (TryGetPartLength(partUri, out long declaredLength) &&
                declaredLength > maximumBytes - currentBytes) {
                throw new OfficePackageSignatureResourceLimitException(
                    "OPC signature inspection exceeds the configured aggregate digest-byte limit.");
            }
        }

        private static long AddDigestWorkBytes(long current, long additional, long maximum) {
            if (additional < 0 || current < 0 || additional > maximum - current) {
                throw new OfficePackageSignatureResourceLimitException(
                    "OPC signature inspection exceeds the configured aggregate digest-byte limit.");
            }
            return current + additional;
        }

        private Dictionary<string, string> ReadContentTypes(long maxPartBytes) {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (!_entries.ContainsKey("/[Content_Types].xml")) return result;

            byte[] contentTypeBytes = ReadPart("/[Content_Types].xml", maxPartBytes);
            using var contentTypeStream = new MemoryStream(contentTypeBytes, writable: false);
            XDocument document = LoadXDocument(contentTypeStream, maxPartBytes);
            XNamespace contentTypes = "http://schemas.openxmlformats.org/package/2006/content-types";
            var defaults = document.Root?
                .Elements(contentTypes + "Default")
                .Where(element => element.Attribute("Extension") != null && element.Attribute("ContentType") != null)
                .ToDictionary(
                    element => ((string)element.Attribute("Extension")!).TrimStart('.'),
                    element => (string)element.Attribute("ContentType")!,
                    StringComparer.OrdinalIgnoreCase)
                ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var overrides = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (XElement element in document.Root?.Elements(contentTypes + "Override") ?? Enumerable.Empty<XElement>()) {
                string partUri = NormalizePartUri((string?)element.Attribute("PartName") ?? string.Empty);
                string? contentType = (string?)element.Attribute("ContentType");
                if (partUri.Length == 0 || string.IsNullOrWhiteSpace(contentType) || overrides.ContainsKey(partUri)) continue;
                overrides.Add(partUri, contentType!);
            }

            foreach (string partUri in _entries.Keys) {
                if (partUri.Equals("/[Content_Types].xml", StringComparison.OrdinalIgnoreCase)) continue;
                overrides.TryGetValue(partUri, out string? contentType);
                if (contentType == null) {
                    string extension = Path.GetExtension(partUri).TrimStart('.');
                    defaults.TryGetValue(extension, out contentType);
                }
                if (!string.IsNullOrWhiteSpace(contentType)) result[partUri] = contentType!;
            }
            return result;
        }

        private static XmlDocument LoadXml(byte[] bytes, long maxCharacters) {
            using var stream = new MemoryStream(bytes, writable: false);
            using XmlReader reader = XmlReader.Create(stream, SafeXmlReaderSettings(maxCharacters));
            XmlDocument document = CreateXmlDocument();
            document.Load(reader);
            return document;
        }

        private static XDocument LoadXDocument(Stream stream, long maxCharacters) {
            using XmlReader reader = XmlReader.Create(stream, SafeXmlReaderSettings(maxCharacters));
            return XDocument.Load(reader, LoadOptions.PreserveWhitespace);
        }

        private static XmlDocument CreateXmlDocument() => new() {
            PreserveWhitespace = true,
            XmlResolver = null
        };

        private static XmlReaderSettings SafeXmlReaderSettings(long maxCharacters = 64L * 1024 * 1024) => new() {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            MaxCharactersInDocument = maxCharacters
        };

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

        private static bool FixedTimeEquals(byte[] left, byte[] right) {
            int difference = left.Length ^ right.Length;
            int length = Math.Min(left.Length, right.Length);
            for (int index = 0; index < length; index++) difference |= left[index] ^ right[index];
            return difference == 0;
        }

        private static void CopyBounded(Stream source, Stream destination, long maxBytes, string partUri) {
            byte[] buffer = new byte[81920];
            long total = 0;
            while (true) {
                int read = source.Read(buffer, 0, buffer.Length);
                if (read == 0) break;
                total += read;
                if (total > maxBytes) {
                    throw new InvalidDataException("The OPC package part '" + partUri + "' exceeds the " + maxBytes + " byte limit.");
                }
                destination.Write(buffer, 0, read);
            }
        }
    }

    internal readonly struct OfficePackageDigestResult {
        private OfficePackageDigestResult(
            OfficePackageSignatureValidationState status,
            string detail,
            long digestWorkBytes) {
            Status = status;
            Detail = detail;
            DigestWorkBytes = digestWorkBytes;
        }

        internal OfficePackageSignatureValidationState Status { get; }
        internal string Detail { get; }
        internal long DigestWorkBytes { get; }

        internal static OfficePackageDigestResult Passed(string detail, long hashedBytes = 0) =>
            new(OfficePackageSignatureValidationState.Passed, detail, hashedBytes);

        internal static OfficePackageDigestResult Failed(string detail, long hashedBytes = 0) =>
            new(OfficePackageSignatureValidationState.Failed, detail, hashedBytes);

        internal static OfficePackageDigestResult Unsupported(string detail, long digestWorkBytes = 0) =>
            new(OfficePackageSignatureValidationState.Unsupported, detail, digestWorkBytes);

        internal static OfficePackageDigestResult NotChecked(string detail) =>
            new(OfficePackageSignatureValidationState.NotChecked, detail, 0);
    }
}
