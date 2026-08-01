#nullable enable
using System.IO.Compression;
using System.Security.Cryptography;
using System.Security.Cryptography.Xml;
using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO.Word {
    /// <summary>Bounded OPC archive reader and transform-aware signature digest engine.</summary>
    internal sealed class OfficePackageSignatureArchive : IDisposable {
        internal const string RelationshipTransformAlgorithm = "http://schemas.openxmlformats.org/package/2006/RelationshipTransform";
        internal const string CanonicalXmlAlgorithm = SignedXml.XmlDsigC14NTransformUrl;
        internal const string CanonicalXmlWithCommentsAlgorithm = SignedXml.XmlDsigC14NWithCommentsTransformUrl;
        internal const string RelationshipsContentType = "application/vnd.openxmlformats-package.relationships+xml";

        private readonly MemoryStream _stream;
        private readonly ZipArchive _archive;
        private readonly Dictionary<string, ZipArchiveEntry> _entries;
        private readonly Dictionary<string, string> _contentTypes;

        internal OfficePackageSignatureArchive(byte[] packageBytes, int maxParts = 10000) {
            if (packageBytes == null) throw new ArgumentNullException(nameof(packageBytes));
            if (maxParts <= 0) throw new ArgumentOutOfRangeException(nameof(maxParts));

            _stream = new MemoryStream(packageBytes, writable: false);
            _archive = new ZipArchive(_stream, ZipArchiveMode.Read, leaveOpen: false);
            if (_archive.Entries.Count > maxParts) {
                throw new InvalidDataException("The OPC package contains more than " + maxParts + " ZIP entries.");
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

            _contentTypes = ReadContentTypes();
        }

        internal IReadOnlyList<string> PartUris => _entries.Keys
            .Where(uri => !uri.Equals("/[Content_Types].xml", StringComparison.OrdinalIgnoreCase))
            .OrderBy(uri => uri, StringComparer.OrdinalIgnoreCase)
            .ToArray();

        internal bool ContainsPart(string partUri) => _entries.ContainsKey(NormalizePartUri(partUri));

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

        internal OfficePackageDigestResult VerifyReference(XElement reference, long maxPartBytes) {
            XNamespace ds = SignedXml.XmlDsigNamespaceUrl;
            string? uri = ((string?)reference.Attribute("URI"))?.Trim();
            string? targetPartUri = NormalizeReferencePartUri(uri);
            if (targetPartUri == null) {
                return OfficePackageDigestResult.NotChecked("Reference '" + (uri ?? string.Empty) + "' is not an OPC package-part reference.");
            }
            if (!ContainsPart(targetPartUri)) {
                return OfficePackageDigestResult.Failed("Digest verification failed because package part " + targetPartUri + " does not exist.");
            }
            string? declaredContentType;
            try {
                declaredContentType = GetDeclaredContentType(uri);
            } catch (UriFormatException exception) {
                return OfficePackageDigestResult.Failed("Digest verification failed because Reference URI content-type encoding is invalid: " + exception.Message);
            }
            if (declaredContentType != null) {
                if (!TryGetContentType(targetPartUri, out string actualContentType)) {
                    return OfficePackageDigestResult.Failed("Digest verification failed because package part " + targetPartUri + " has no resolved OPC content type.");
                }
                if (!string.Equals(declaredContentType, actualContentType, StringComparison.OrdinalIgnoreCase)) {
                    return OfficePackageDigestResult.Failed(
                        "Digest verification failed because package part " + targetPartUri +
                        " declares signed content type '" + declaredContentType +
                        "' but the package resolves it as '" + actualContentType + "'.");
                }
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
            try {
                input = ApplyTransforms(targetPartUri, reference, maxPartBytes);
            } catch (NotSupportedException exception) {
                return OfficePackageDigestResult.Unsupported(exception.Message);
            } catch (Exception exception) when (exception is IOException or InvalidDataException or XmlException or CryptographicException) {
                return OfficePackageDigestResult.Failed("Digest verification for " + targetPartUri + " failed while applying transforms: " + exception.Message);
            }

            byte[] actual;
            using (HashAlgorithm hash = hashFactory()) {
                actual = hash.ComputeHash(input);
            }

            return FixedTimeEquals(actual, expected)
                ? OfficePackageDigestResult.Passed("Digest verification passed for " + targetPartUri + ".")
                : OfficePackageDigestResult.Failed("Digest verification failed for " + targetPartUri + ".");
        }

        internal string ComputeDigestValue(XElement reference, long maxPartBytes) {
            XNamespace ds = SignedXml.XmlDsigNamespaceUrl;
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
            if (declaredContentType != null &&
                (!TryGetContentType(targetPartUri, out string actualContentType) ||
                 !string.Equals(declaredContentType, actualContentType, StringComparison.OrdinalIgnoreCase))) {
                throw new InvalidDataException("The OPC signature Reference content type does not match the package content type for " + targetPartUri + ".");
            }
            byte[] input = ApplyTransforms(targetPartUri, reference, maxPartBytes);
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

        internal static byte[] Canonicalize(XmlDocument document, bool includeComments = false) {
            Transform transform = includeComments
                ? new XmlDsigC14NWithCommentsTransform()
                : new XmlDsigC14NTransform();
            transform.LoadInput(document);
            using Stream output = (Stream)transform.GetOutput(typeof(Stream));
            using var memory = new MemoryStream();
            output.CopyTo(memory);
            return memory.ToArray();
        }

        private byte[] ApplyTransforms(string targetPartUri, XElement reference, long maxPartBytes) {
            XNamespace ds = SignedXml.XmlDsigNamespaceUrl;
            List<XElement> transforms = reference
                .Element(ds + "Transforms")?
                .Elements(ds + "Transform")
                .ToList() ?? new List<XElement>();

            if (transforms.Count == 0) return ReadPart(targetPartUri, maxPartBytes);

            byte[]? currentBytes = null;
            XmlDocument? currentXml = null;
            foreach (XElement transform in transforms) {
                string algorithm = ((string?)transform.Attribute("Algorithm"))?.Trim() ?? string.Empty;
                if (algorithm.Equals(RelationshipTransformAlgorithm, StringComparison.Ordinal)) {
                    if (currentBytes != null || currentXml != null) {
                        throw new NotSupportedException("RelationshipTransform must be the first transform for " + targetPartUri + ".");
                    }
                    currentXml = ApplyRelationshipTransform(targetPartUri, transform, maxPartBytes);
                    continue;
                }

                if (algorithm.Equals(CanonicalXmlAlgorithm, StringComparison.Ordinal) ||
                    algorithm.Equals(CanonicalXmlWithCommentsAlgorithm, StringComparison.Ordinal)) {
                    if (currentXml == null) {
                        currentXml = LoadXml(currentBytes ?? ReadPart(targetPartUri, maxPartBytes), maxPartBytes);
                    }
                    currentBytes = Canonicalize(
                        currentXml,
                        algorithm.Equals(CanonicalXmlWithCommentsAlgorithm, StringComparison.Ordinal));
                    currentXml = null;
                    continue;
                }

                throw new NotSupportedException("Digest verification for " + targetPartUri + " does not support transform " + algorithm + ".");
            }

            if (currentXml != null) {
                throw new NotSupportedException("RelationshipTransform for " + targetPartUri + " must be followed by an XML canonicalization transform.");
            }
            return currentBytes ?? ReadPart(targetPartUri, maxPartBytes);
        }

        private XmlDocument ApplyRelationshipTransform(string targetPartUri, XElement transform, long maxPartBytes) {
            XmlDocument source = LoadXml(ReadPart(targetPartUri, maxPartBytes), maxPartBytes);
            const string relationshipNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";
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

        private Dictionary<string, string> ReadContentTypes() {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (!_entries.TryGetValue("/[Content_Types].xml", out ZipArchiveEntry? entry)) return result;

            XDocument document;
            using (Stream stream = entry.Open()) {
                document = LoadXDocument(stream);
            }
            XNamespace contentTypes = "http://schemas.openxmlformats.org/package/2006/content-types";
            var defaults = document.Root?
                .Elements(contentTypes + "Default")
                .Where(element => element.Attribute("Extension") != null && element.Attribute("ContentType") != null)
                .ToDictionary(
                    element => ((string)element.Attribute("Extension")!).TrimStart('.'),
                    element => (string)element.Attribute("ContentType")!,
                    StringComparer.OrdinalIgnoreCase)
                ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            foreach (string partUri in _entries.Keys) {
                if (partUri.Equals("/[Content_Types].xml", StringComparison.OrdinalIgnoreCase)) continue;
                string? contentType = document.Root?
                    .Elements(contentTypes + "Override")
                    .Where(element => string.Equals(
                        NormalizePartUri((string?)element.Attribute("PartName") ?? string.Empty),
                        partUri,
                        StringComparison.OrdinalIgnoreCase))
                    .Select(element => (string?)element.Attribute("ContentType"))
                    .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));
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

        private static XDocument LoadXDocument(Stream stream) {
            using XmlReader reader = XmlReader.Create(stream, SafeXmlReaderSettings());
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
                case SignedXml.XmlDsigSHA1Url:
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
            OfficePackageSignatureDigestVerificationStatus status,
            string detail) {
            Status = status;
            Detail = detail;
        }

        internal OfficePackageSignatureDigestVerificationStatus Status { get; }
        internal string Detail { get; }

        internal static OfficePackageDigestResult Passed(string detail) =>
            new(OfficePackageSignatureDigestVerificationStatus.Passed, detail);

        internal static OfficePackageDigestResult Failed(string detail) =>
            new(OfficePackageSignatureDigestVerificationStatus.Failed, detail);

        internal static OfficePackageDigestResult Unsupported(string detail) =>
            new(OfficePackageSignatureDigestVerificationStatus.Unsupported, detail);

        internal static OfficePackageDigestResult NotChecked(string detail) =>
            new(OfficePackageSignatureDigestVerificationStatus.NotChecked, detail);
    }
}
