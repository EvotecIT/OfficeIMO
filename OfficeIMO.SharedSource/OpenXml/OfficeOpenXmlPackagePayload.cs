using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using OfficeIMO.Drawing.Internal;
using System.Security.Cryptography;

namespace OfficeIMO.OpenXml.Internal {
    internal sealed class OfficeOpenXmlPayloadHandle {
        internal OfficeOpenXmlPayloadHandle(
            OpenXmlPartContainer owner,
            OpenXmlPart part,
            string relationshipId,
            OfficeEmbeddedPayloadKind kind) {
            Owner = owner;
            Part = part;
            RelationshipId = relationshipId;
            Kind = kind;
        }

        internal OpenXmlPartContainer Owner { get; }

        internal OpenXmlPart Part { get; }

        internal string RelationshipId { get; }

        internal OfficeEmbeddedPayloadKind Kind { get; }

        internal string OwnerPartUri => Owner is OpenXmlPart ownerPart ? ownerPart.Uri.ToString() : "/";

        internal string Id => OwnerPartUri + "#" + RelationshipId;
    }

    /// <summary>Shared package-part inventory and bounded stream operations for Office Open XML formats.</summary>
    internal static class OfficeOpenXmlPackagePayload {
        internal static IReadOnlyList<OfficeOpenXmlPayloadHandle> FindEmbeddedPayloads(OpenXmlPackage package) {
            if (package == null) throw new ArgumentNullException(nameof(package));

            var result = new List<OfficeOpenXmlPayloadHandle>();
            var visited = new HashSet<OpenXmlPart>();
            var queue = new Queue<OpenXmlPartContainer>();
            queue.Enqueue(package);
            while (queue.Count > 0) {
                OpenXmlPartContainer owner = queue.Dequeue();
                foreach (IdPartPair pair in owner.Parts) {
                    OpenXmlPart part = pair.OpenXmlPart;
                    if (TryClassify(part, out OfficeEmbeddedPayloadKind kind)) {
                        result.Add(new OfficeOpenXmlPayloadHandle(owner, part, pair.RelationshipId, kind));
                    }

                    if (visited.Add(part)) {
                        queue.Enqueue(part);
                    }
                }
            }

            return result
                .OrderBy(static handle => handle.OwnerPartUri, StringComparer.OrdinalIgnoreCase)
                .ThenBy(static handle => handle.RelationshipId, StringComparer.Ordinal)
                .ToArray();
        }

        internal static OfficeOpenXmlPayloadHandle FindEmbeddedPayload(OpenXmlPackage package, string id) {
            if (string.IsNullOrWhiteSpace(id)) throw new ArgumentException("Payload id cannot be empty.", nameof(id));
            return FindEmbeddedPayloads(package)
                .FirstOrDefault(handle => string.Equals(handle.Id, id, StringComparison.Ordinal))
                ?? throw new KeyNotFoundException($"Embedded package payload '{id}' was not found.");
        }

        internal static OfficeEmbeddedPayloadInfo CreateInfo(OfficeOpenXmlPayloadHandle handle, bool includeSha256) {
            if (handle == null) throw new ArgumentNullException(nameof(handle));
            long length = GetLength(handle.Part);
            string? sha256 = includeSha256 ? ComputeSha256(handle.Part) : null;
            string partUri = handle.Part.Uri.ToString();
            string fileName = Path.GetFileName(partUri.Replace('/', Path.DirectorySeparatorChar));
            return new OfficeEmbeddedPayloadInfo(
                handle.Id,
                handle.Kind,
                handle.OwnerPartUri,
                handle.RelationshipId,
                partUri,
                handle.Part.ContentType,
                fileName,
                length,
                sha256);
        }

        internal static byte[] ReadBytes(OpenXmlPart part, long? maxBytes = null) {
            if (part == null) throw new ArgumentNullException(nameof(part));
            using Stream stream = part.GetStream(FileMode.Open, FileAccess.Read);
            return OfficeStreamReader.ReadAllBytes(stream, maxBytes);
        }

        internal static void SaveBytes(OpenXmlPart part, string filePath, long? maxBytes = null) {
            if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException("Destination path cannot be empty.", nameof(filePath));
            OfficeFileCommit.WriteAllBytes(filePath, ReadBytes(part, maxBytes));
        }

        internal static void ReplaceBytes(OpenXmlPart part, byte[] data) {
            if (part == null) throw new ArgumentNullException(nameof(part));
            if (data == null || data.Length == 0) throw new ArgumentException("Payload data cannot be empty.", nameof(data));
            using var source = new MemoryStream(data, writable: false);
            part.FeedData(source);
        }

        internal static void RemovePart(OfficeOpenXmlPayloadHandle handle) {
            if (handle == null) throw new ArgumentNullException(nameof(handle));
            RemoveKnownRelationshipReferences(handle);
            handle.Owner.DeletePart(handle.Part);
        }

        private static void RemoveKnownRelationshipReferences(OfficeOpenXmlPayloadHandle handle) {
            if (handle.Owner is not OpenXmlPart ownerPart || ownerPart.RootElement == null) {
                return;
            }

            OpenXmlElement[] references = ownerPart.RootElement
                .Descendants()
                .Where(element => element.GetAttributes().Any(attribute =>
                    string.Equals(attribute.NamespaceUri, "http://schemas.openxmlformats.org/officeDocument/2006/relationships", StringComparison.Ordinal)
                    && string.Equals(attribute.Value, handle.RelationshipId, StringComparison.Ordinal)))
                .ToArray();
            foreach (OpenXmlElement reference in references) {
                OpenXmlElement? target = GetKnownPayloadReferenceRoot(reference);
                if (target == null) {
                    throw new InvalidOperationException(
                        $"The payload relationship '{handle.RelationshipId}' is referenced by unsupported element '{reference.LocalName}'. Replace or extract the payload instead of removing it.");
                }

                target.Remove();
            }

            foreach (OpenXmlElement container in ownerPart.RootElement.Descendants()
                .Where(static element => IsEmptyPayloadContainer(element))
                .ToArray()) {
                container.Remove();
            }
        }

        private static OpenXmlElement? GetKnownPayloadReferenceRoot(OpenXmlElement element) {
            if (string.Equals(element.LocalName, "oleObject", StringComparison.OrdinalIgnoreCase)) {
                return element.Ancestors()
                    .FirstOrDefault(ancestor => string.Equals(ancestor.LocalName, "object", StringComparison.OrdinalIgnoreCase))
                    ?? element;
            }

            if (string.Equals(element.LocalName, "object", StringComparison.OrdinalIgnoreCase)
                || string.Equals(element.LocalName, "control", StringComparison.OrdinalIgnoreCase)) {
                return element;
            }

            return null;
        }

        private static bool IsEmptyPayloadContainer(OpenXmlElement element) {
            return !element.ChildElements.Any()
                && (string.Equals(element.LocalName, "oleObjects", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(element.LocalName, "controls", StringComparison.OrdinalIgnoreCase));
        }

        internal static OfficeVbaProjectInfo CreateVbaProjectInfo(VbaProjectPart part, bool includeSha256,
            long maxBytes = OfficeVbaProjectInfo.DefaultMaximumProjectBytes) {
            if (part == null) throw new ArgumentNullException(nameof(part));
            if (maxBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maxBytes));
            byte[] projectBytes = ReadBytes(part, maxBytes);
            string? sha256 = includeSha256 ? ComputeSha256(projectBytes) : null;
            bool hasSignature = part.Parts.Any(pair =>
                pair.OpenXmlPart.RelationshipType.IndexOf("vbaProjectSignature", StringComparison.OrdinalIgnoreCase) >= 0
                || pair.OpenXmlPart.ContentType.IndexOf("vbaProjectSignature", StringComparison.OrdinalIgnoreCase) >= 0);
            return new OfficeVbaProjectInfo(
                part.Uri.ToString(),
                part.ContentType,
                projectBytes.LongLength,
                sha256,
                OfficeVbaProjectInspector.GetModuleNames(projectBytes),
                hasSignature);
        }

        private static bool TryClassify(OpenXmlPart part, out OfficeEmbeddedPayloadKind kind) {
            string uri = part.Uri.ToString();
            string contentType = part.ContentType ?? string.Empty;
            if (part is EmbeddedPackagePart) {
                kind = OfficeEmbeddedPayloadKind.EmbeddedPackage;
                return true;
            }
            if (part is EmbeddedObjectPart || contentType.IndexOf("oleObject", StringComparison.OrdinalIgnoreCase) >= 0) {
                kind = OfficeEmbeddedPayloadKind.OleObject;
                return true;
            }
            if (uri.IndexOf("/activeX/", StringComparison.OrdinalIgnoreCase) >= 0
                || contentType.IndexOf("activeX", StringComparison.OrdinalIgnoreCase) >= 0) {
                kind = OfficeEmbeddedPayloadKind.ActiveX;
                return true;
            }
            if (uri.IndexOf("/embeddings/", StringComparison.OrdinalIgnoreCase) >= 0) {
                kind = OfficeEmbeddedPayloadKind.Other;
                return true;
            }

            kind = default;
            return false;
        }

        private static long GetLength(OpenXmlPart part) {
            using Stream stream = part.GetStream(FileMode.Open, FileAccess.Read);
            if (stream.CanSeek) {
                return stream.Length;
            }

            long length = 0;
            var buffer = new byte[81920];
            int read;
            while ((read = stream.Read(buffer, 0, buffer.Length)) > 0) {
                length = checked(length + read);
            }
            return length;
        }

        private static string ComputeSha256(OpenXmlPart part) {
            using Stream stream = part.GetStream(FileMode.Open, FileAccess.Read);
            using SHA256 sha = SHA256.Create();
            return ToHex(sha.ComputeHash(stream));
        }

        private static string ComputeSha256(byte[] data) {
            using SHA256 sha = SHA256.Create();
            return ToHex(sha.ComputeHash(data));
        }

        private static string ToHex(byte[] bytes) {
            var builder = new System.Text.StringBuilder(bytes.Length * 2);
            foreach (byte value in bytes) {
                builder.Append(value.ToString("x2", System.Globalization.CultureInfo.InvariantCulture));
            }
            return builder.ToString();
        }
    }
}
