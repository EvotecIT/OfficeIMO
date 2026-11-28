using System.Buffers;
using System.Security.Cryptography;
using OfficeIMO.Core.Internal;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.PowerPoint.OpenXml;

namespace OfficeIMO.PowerPoint;

public sealed partial class PowerPointDocument {
    private const string EmbeddedPackageRelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "package";
    private const string EmbeddedObjectRelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "oleObject";
    private const string VideoDataRelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "video";
    private const string AudioDataRelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "audio";
    private const string MediaDataRelationshipType = "http://schemas.microsoft.com/office/2007/relationships/media";

    // Sessions may read the same borrowed ZIP stream. Keep discovery and each selected read on one cursor at a time.
    private SemaphoreSlim? _payloadReadGate;
    private SemaphoreSlim PayloadReadGate => LazyInitializer.EnsureInitialized(ref _payloadReadGate, static () => new SemaphoreSlim(1, 1))!;

    private List<OfficeEmbeddedPayloadInfo>? GetPackagePayloads(CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (!_sdkDeferred || IsSdkLoaded) return null;
        SemaphoreSlim gate = PayloadReadGate;
        gate.Wait(cancellationToken);
        try { return CollectPackagePayloads(cancellationToken); }
        finally { gate.Release(); }
    }

    private List<OfficeEmbeddedPayloadInfo>? CollectPackagePayloads(CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (ReadablePackage is not { } package) return null;
        var found = new List<OfficeEmbeddedPayloadInfo>();
        if (!TryVisitPackageGraph(package, cancellationToken, relationship => {
                OfficeEmbeddedPayloadKind? kind = ClassifyPayload(relationship.Type, relationship.Part, relationship.ContentType);
                if (kind == null) return;
                string owner = relationship.Owner ?? "/";
                string fileName = Path.GetFileName(relationship.Part.Replace('/', Path.DirectorySeparatorChar));
                found.Add(new OfficeEmbeddedPayloadInfo(owner + "#" + relationship.Id, kind.Value, owner,
                    relationship.Id, relationship.Part, relationship.ContentType, fileName, -1, sha256: null));
            })) return null;

        found.Sort(static (left, right) => {
            int owner = StringComparer.OrdinalIgnoreCase.Compare(left.OwnerPartUri, right.OwnerPartUri);
            return owner != 0 ? owner : StringComparer.Ordinal.Compare(left.RelationshipId, right.RelationshipId);
        });
        return found;
    }

    private readonly record struct PackageGraphRelationship(string? Owner, string Id, string Type, string Part, string ContentType);

    private static bool TryVisitPackageGraph(OfficeOpcEditablePackage package, CancellationToken cancellationToken,
        Action<PackageGraphRelationship> visit, string? root = null) {
        if (package.PartNames.Any(name => name.EndsWith(".piece", StringComparison.OrdinalIgnoreCase) &&
            OfficeOpcPieceAssembler.HasPieceSuffix(name))) return false;
        return TryVisitPackageGraph(owner => ReadEditablePackageRelationships(package, owner), package.GetPartName,
            package.GetContentType, cancellationToken, visit, root);
    }

    private static bool TryVisitPackageGraph(OfficeOpcPackage package, CancellationToken cancellationToken,
        Action<PackageGraphRelationship> visit) {
        if (package.HasInterleavedParts) return false;
        try {
            return TryVisitPackageGraph(package.GetRelationships, package.GetPartName, package.GetContentType,
                cancellationToken, visit);
        } catch (Exception exception) when (exception is InvalidDataException or IOException or System.Xml.XmlException
            or ArgumentException or NotSupportedException or InvalidOperationException or OverflowException) {
            // The existing SDK opener decides how to report or repair malformed relationships.
            return false;
        }
    }

    private static bool TryVisitPackageGraph(Func<string?, IReadOnlyList<OfficeOpcRelationship>?> getRelationships,
        Func<string, string?> getPartName, Func<string, string?> getContentType, CancellationToken cancellationToken,
        Action<PackageGraphRelationship> visit, string? root = null) {
        var visited = new Dictionary<string, string>(OfficeAsciiIgnoreCaseComparer.Instance);
        var pending = new Queue<string?>();
        pending.Enqueue(root);
        while (pending.Count > 0) {
            cancellationToken.ThrowIfCancellationRequested();
            string? owner = pending.Dequeue();
            if (getRelationships(owner) is not { } relationships) return false;
            if (relationships.Count == 0) continue;
            var targets = new HashSet<string>(OfficeAsciiIgnoreCaseComparer.Instance);
            for (int index = 0; index < relationships.Count; index++) {
                cancellationToken.ThrowIfCancellationRequested();
                OfficeOpcRelationship relationship = relationships[index];
                if (relationship.Id.Length == 0 || relationship.Type.Length == 0) return false;
                if (relationship.IsExternal || IsReferenceRelationship(relationship.Type)) continue;
                if (relationship.TargetPartName is not { } target || getPartName(target) is not { } part) return false;
                string identity = OfficeOpcPackage.GetPartIdentity(part);
                if (!targets.Add(identity)) return false;
                if (visited.TryGetValue(identity, out string? previousType)) {
                    if (previousType != relationship.Type) return false;
                } else {
                    visited.Add(identity, relationship.Type);
                    pending.Enqueue(part);
                }

                string? contentType = getContentType(part);
                if (contentType == null || !CanFollowPackageGraphPart(relationship.Type, contentType)) return false;
                visit(new PackageGraphRelationship(owner, relationship.Id, relationship.Type, part, contentType));
            }
        }
        return true;
    }

    private static IReadOnlyList<OfficeOpcRelationship>? ReadEditablePackageRelationships(OfficeOpcEditablePackage package, string? owner) {
        string relationshipsPart = OfficeOpcPartNames.GetRelationshipsPartName(owner);
        if (!package.PartExists(relationshipsPart)) return Array.Empty<OfficeOpcRelationship>();
        OfficeXmlTree tree = package.GetXml(relationshipsPart);
        const int ns = OfficeXmlNamespaces.PackageRelationships;
        if (tree.Root == OfficeXmlTree.NoNode || !tree.IsElement(tree.Root, ns, "Relationships")) return null;
        var relationships = new List<OfficeOpcRelationship>();
        var ids = new HashSet<string>(StringComparer.Ordinal);
        string? directory = OfficeOpcPartNames.GetRelationshipSourceDirectory(relationshipsPart);
        for (int node = tree.FindChild(tree.Root, ns, "Relationship"); node != OfficeXmlTree.NoNode;
             node = tree.FindNextSibling(node, ns, "Relationship")) {
            string? type = tree.GetAttribute(node, OfficeXmlNamespaces.None, "Type");
            string? id = tree.GetAttribute(node, OfficeXmlNamespaces.None, "Id");
            string? target = tree.GetAttribute(node, OfficeXmlNamespaces.None, "Target");
            if (string.IsNullOrEmpty(type) || string.IsNullOrEmpty(id) || target == null || !ids.Add(id!)) return null;
            string? mode = tree.GetAttribute(node, OfficeXmlNamespaces.None, "TargetMode");
            if (mode != null && mode != "Internal" && mode != "External") return null;
            relationships.Add(new OfficeOpcRelationship(directory, id!, OfficeXmlNamespaces.NormalizeRelationshipType(type!), target, mode == "External"));
        }
        return relationships;
    }

    private static bool IsReferenceRelationship(string type) =>
        type == PowerPointPartRef.HyperlinkRelationshipType || type == AudioDataRelationshipType ||
        type == VideoDataRelationshipType || type == MediaDataRelationshipType;

    private static bool CanFollowPackageGraphPart(string relationship, string contentType) {
        string? expected = relationship switch {
            PowerPointPartRef.SlideRelationshipType => PowerPointPackageModel.SlideContentType,
            PowerPointPartRef.SlideLayoutRelationshipType => PowerPointPackageModel.SlideLayoutContentType,
            PowerPointPartRef.SlideMasterRelationshipType => PowerPointPackageModel.SlideMasterContentType,
            PowerPointPartRef.ThemeRelationshipType => PowerPointPackageModel.ThemeContentType,
            NotesSlideRelationshipType => NotesSlideContentType,
            VbaProjectRelationshipType => VbaProjectContentType,
            _ => null
        };
        return expected == null || string.Equals(expected, contentType, StringComparison.OrdinalIgnoreCase);
    }

    private static OfficeEmbeddedPayloadKind? ClassifyPayload(string relationship, string partUri, string contentType) {
        if (relationship == EmbeddedPackageRelationshipType) return OfficeEmbeddedPayloadKind.EmbeddedPackage;
        if (relationship == EmbeddedObjectRelationshipType ||
            contentType.IndexOf("oleObject", StringComparison.OrdinalIgnoreCase) >= 0) return OfficeEmbeddedPayloadKind.OleObject;
        if (partUri.IndexOf("/activeX/", StringComparison.OrdinalIgnoreCase) >= 0 ||
            contentType.IndexOf("activeX", StringComparison.OrdinalIgnoreCase) >= 0) return OfficeEmbeddedPayloadKind.ActiveX;
        return partUri.IndexOf("/embeddings/", StringComparison.OrdinalIgnoreCase) >= 0 ? OfficeEmbeddedPayloadKind.Other : null;
    }

    private OfficeEmbeddedPayloadSource CreatePackagePayloadSource(OfficeEmbeddedPayloadInfo info) =>
        new(info, token => OpenPackagePayloadStream(info.PartUri, token), executeCancelableRead: (action, cancellationToken) => {
            SemaphoreSlim gate = PayloadReadGate;
            gate.Wait(cancellationToken);
            try { action(); }
            finally { gate.Release(); }
        });

    private Stream OpenPackagePayloadStream(string part, CancellationToken cancellationToken) {
        ThrowIfDisposed();
        cancellationToken.ThrowIfCancellationRequested();
        if (ReadablePackage is { } package) return package.OpenPartContentStream(part, cancellationToken);
        return PowerPointPartRef.Package(this, part).OpenRead();
    }

    private OfficeEmbeddedPayloadInfo ReadPackagePayloadInfo(OfficeEmbeddedPayloadInfo info, bool includeSha256) {
        SemaphoreSlim gate = PayloadReadGate;
        gate.Wait();
        try { return ReadPackagePayloadDetails(info, includeSha256); }
        finally { gate.Release(); }
    }

    private OfficeEmbeddedPayloadInfo ReadPackagePayloadDetails(OfficeEmbeddedPayloadInfo info, bool includeSha256) {
        long length;
        string? hash = null;
        if (!includeSha256 && ReadablePackage is { } package && package.TryGetBinaryLength(info.PartUri, out length)) {
            return new OfficeEmbeddedPayloadInfo(info.Id, info.Kind, info.OwnerPartUri, info.RelationshipId, info.PartUri,
                info.ContentType, info.SuggestedFileName, length, hash);
        }

        using Stream stream = OpenPackagePayloadStream(info.PartUri, default);
        using IncrementalHash? sha = includeSha256 ? IncrementalHash.CreateHash(HashAlgorithmName.SHA256) : null;
        length = 0;
        byte[] buffer = ArrayPool<byte>.Shared.Rent(81920);
        try {
            int read;
            while ((read = stream.Read(buffer, 0, buffer.Length)) > 0) {
                length = checked(length + read);
                sha?.AppendData(buffer, 0, read);
            }
        } finally {
            ArrayPool<byte>.Shared.Return(buffer);
        }
        if (sha != null) hash = BitConverter.ToString(sha.GetHashAndReset()).Replace("-", string.Empty).ToLowerInvariant();
        return new OfficeEmbeddedPayloadInfo(info.Id, info.Kind, info.OwnerPartUri, info.RelationshipId, info.PartUri,
            info.ContentType, info.SuggestedFileName, length, hash);
    }
}
