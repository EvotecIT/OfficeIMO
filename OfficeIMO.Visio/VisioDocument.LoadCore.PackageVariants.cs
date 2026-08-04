using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using OfficeIMO.Drawing.Internal;

namespace OfficeIMO.Visio {
    public partial class VisioDocument {
        private const int MaximumVbaSubtreeParts = 256;
        private const long MaximumVbaSubtreeBytes = 128L * 1024L * 1024L;

        private sealed class PreservedVbaRelationship {
            internal PreservedVbaRelationship(string id, string type,
                Uri targetUri, TargetMode targetMode) {
                Id = id;
                Type = type;
                TargetUri = targetUri;
                TargetMode = targetMode;
            }
            internal string Id { get; }
            internal string Type { get; }
            internal Uri TargetUri { get; }
            internal TargetMode TargetMode { get; }
        }

        private sealed class PreservedVbaPart {
            internal PreservedVbaPart(Uri uri, string contentType, byte[] data,
                IReadOnlyList<PreservedVbaRelationship> relationships) {
                Uri = uri;
                ContentType = contentType;
                Data = data;
                Relationships = relationships;
            }
            internal Uri Uri { get; }
            internal string ContentType { get; }
            internal byte[] Data { get; }
            internal IReadOnlyList<PreservedVbaRelationship> Relationships { get; }
        }

        private static void LoadVbaProject(PackagePart documentPart, VisioDocument document) {
            PackageRelationship[] relationships = documentPart
                .GetRelationshipsByType(VbaProjectRelationshipType).ToArray();
            if (relationships.Length > 1)
                throw new InvalidDataException("A Visio document may contain only one VBA project relationship.");
            PackageRelationship? relationship = relationships.SingleOrDefault();
            if (relationship == null) return;
            if (relationship.TargetMode != TargetMode.Internal) {
                throw new InvalidDataException("A Visio VBA project relationship must target an internal package part.");
            }

            Uri partUri = PackUriHelper.ResolvePartUri(documentPart.Uri,
                relationship.TargetUri);
            Package package = documentPart.Package;
            if (!package.PartExists(partUri)) {
                throw new InvalidDataException("The Visio VBA project relationship targets a missing package part.");
            }

            CaptureVbaSubtree(package, partUri, document);
            PreservedVbaPart root = document._preservedVbaParts[
                partUri.OriginalString];
            document._vbaProjectPartUri = partUri;
            document._vbaProjectBytes = root.Data;
            document._vbaProjectContentType = root.ContentType;
        }

        private static void CaptureVbaSubtree(Package package, Uri rootUri,
            VisioDocument document) {
            var pending = new Queue<Uri>();
            var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            pending.Enqueue(rootUri);
            long totalBytes = 0L;
            while (pending.Count > 0) {
                Uri uri = pending.Dequeue();
                if (!visited.Add(uri.OriginalString)) continue;
                if (visited.Count > MaximumVbaSubtreeParts)
                    throw new InvalidDataException(
                        "The Visio VBA relationship subtree contains too many package parts.");
                if (!package.PartExists(uri))
                    throw new InvalidDataException(
                        $"The Visio VBA relationship subtree targets missing part '{uri}'.");
                PackagePart part = package.GetPart(uri);
                byte[] bytes;
                using (Stream stream = part.GetStream(FileMode.Open,
                           FileAccess.Read)) {
                    bytes = OfficeStreamReader.ReadAllBytes(stream,
                        64L * 1024L * 1024L);
                }
                totalBytes = checked(totalBytes + bytes.LongLength);
                if (totalBytes > MaximumVbaSubtreeBytes)
                    throw new InvalidDataException(
                        "The Visio VBA relationship subtree exceeds the supported preservation size.");
                List<PreservedVbaRelationship> relationships = part
                    .GetRelationships().Select(item =>
                        new PreservedVbaRelationship(item.Id,
                            item.RelationshipType, item.TargetUri,
                            item.TargetMode)).ToList();
                document._preservedVbaParts[uri.OriginalString] =
                    new PreservedVbaPart(uri, part.ContentType, bytes,
                        relationships);
                foreach (PreservedVbaRelationship item in relationships
                             .Where(item => item.TargetMode ==
                                            TargetMode.Internal)) {
                    Uri target = PackUriHelper.ResolvePartUri(uri,
                        item.TargetUri);
                    if (!package.PartExists(target))
                        throw new InvalidDataException(
                            $"The Visio VBA relationship subtree targets missing part '{target}'.");
                    pending.Enqueue(target);
                }
            }
        }
    }
}
