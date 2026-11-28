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
        // Interleaved OPC stores physical pieces where an atomic relationship part would otherwise be.
        if (package.PartNames.Any(name => name.EndsWith(".piece", StringComparison.OrdinalIgnoreCase) &&
            OfficeOpcPieceAssembler.HasPieceSuffix(name))) return null;
        var found = new List<OfficeEmbeddedPayloadInfo>();
        var visited = new Dictionary<string, string>(OfficeAsciiIgnoreCaseComparer.Instance);
        var pending = new Queue<string?>();
        pending.Enqueue(null);
        while (pending.Count > 0) {
            cancellationToken.ThrowIfCancellationRequested();
            string? owner = pending.Dequeue();
            string relationshipsPart = OfficeOpcPartNames.GetRelationshipsPartName(owner);
            if (!package.PartExists(relationshipsPart)) continue;
            OfficeXmlTree tree = package.GetXml(relationshipsPart);
            const int ns = OfficeXmlNamespaces.PackageRelationships;
            if (tree.Root == OfficeXmlTree.NoNode || !tree.IsElement(tree.Root, ns, "Relationships")) return null;
            var targets = new HashSet<string>(OfficeAsciiIgnoreCaseComparer.Instance);
            var ids = new HashSet<string>(StringComparer.Ordinal);
            for (int node = tree.FindChild(tree.Root, ns, "Relationship"); node != OfficeXmlTree.NoNode;
                 node = tree.FindNextSibling(node, ns, "Relationship")) {
                cancellationToken.ThrowIfCancellationRequested();
                string? rawType = tree.GetAttribute(node, OfficeXmlNamespaces.None, "Type");
                string? id = tree.GetAttribute(node, OfficeXmlNamespaces.None, "Id");
                if (string.IsNullOrEmpty(rawType) || string.IsNullOrEmpty(id) || !ids.Add(id!)) return null;

                string? mode = tree.GetAttribute(node, OfficeXmlNamespaces.None, "TargetMode");
                if (mode == "External") continue;
                if (mode != null && mode != "Internal") return null;
                string type = OfficeXmlNamespaces.NormalizeRelationshipType(rawType!);
                if (IsReferenceRelationship(type)) continue;

                string? target = OfficeOpcPartNames.ResolveRelationshipTarget(relationshipsPart,
                    tree.GetAttribute(node, OfficeXmlNamespaces.None, "Target"));
                if (target == null || package.GetPartName(target) is not { } part) return null;
                string identity = OfficeOpcPackage.GetPartIdentity(part);
                if (!targets.Add(identity)) return null;
                if (visited.TryGetValue(identity, out string? previousType)) {
                    if (previousType != type) return null;
                } else {
                    visited.Add(identity, type);
                    pending.Enqueue(part);
                }

                string? contentType = package.GetContentType(part);
                if (contentType == null || !CanFollowPayloadGraphPart(type, contentType)) return null;
                OfficeEmbeddedPayloadKind? kind = ClassifyPayload(type, part, contentType);
                if (kind == null) continue;
                string ownerName = owner ?? "/";
                string fileName = Path.GetFileName(part.Replace('/', Path.DirectorySeparatorChar));
                found.Add(new OfficeEmbeddedPayloadInfo(ownerName + "#" + id, kind.Value, ownerName, id!,
                    part, contentType, fileName, -1, sha256: null));
            }
        }

        found.Sort(static (left, right) => {
            int owner = StringComparer.OrdinalIgnoreCase.Compare(left.OwnerPartUri, right.OwnerPartUri);
            return owner != 0 ? owner : StringComparer.Ordinal.Compare(left.RelationshipId, right.RelationshipId);
        });
        return found;
    }

    private static bool IsReferenceRelationship(string type) =>
        type == PowerPointPartRef.HyperlinkRelationshipType || type == AudioDataRelationshipType ||
        type == VideoDataRelationshipType || type == MediaDataRelationshipType;

    private static bool CanFollowPayloadGraphPart(string relationship, string contentType) {
        string? expected = relationship switch {
            PowerPointPartRef.SlideRelationshipType => PowerPointPackageModel.SlideContentType,
            PowerPointPartRef.SlideLayoutRelationshipType => PowerPointPackageModel.SlideLayoutContentType,
            PowerPointPartRef.SlideMasterRelationshipType => PowerPointPackageModel.SlideMasterContentType,
            PowerPointPartRef.ThemeRelationshipType => PowerPointPackageModel.ThemeContentType,
            NotesSlideRelationshipType => NotesSlideContentType,
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
