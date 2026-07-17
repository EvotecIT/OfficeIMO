namespace OfficeIMO.OneNote;

/// <summary>Materializes revision chains for section, page, history, and TOC object spaces.</summary>
internal sealed class OneNoteObjectSpaceMaterializer {
    private static readonly Guid HeaderCellObjectSpaceId = new Guid("111E4CF3-7FEF-4087-AF6A-B9544ACD334D");
    private readonly Dictionary<OneNoteExtendedGuid, OneNoteRevisionManifest> _revisions;
    private readonly Dictionary<OneNoteExtendedGuid, List<OneNoteRevisionManifest>> _spaces;
    private readonly Dictionary<OneNoteExtendedGuid, List<OneNoteRevisionStoreObject>> _objectsByRevision;
    private readonly Dictionary<string, OneNoteMaterializedObjectSpace> _cache = new Dictionary<string, OneNoteMaterializedObjectSpace>(StringComparer.Ordinal);
    private readonly Dictionary<Guid, OneNoteFileDataStoreObject> _fileData;
    private readonly HashSet<OneNoteExtendedGuid> _mappedObjectIds = new HashSet<OneNoteExtendedGuid>();

    public OneNoteObjectSpaceMaterializer(OneNoteRevisionStore store) {
        _revisions = store.Revisions.ToDictionary(revision => revision.Id, revision => revision);
        _spaces = store.Revisions.Where(revision => revision.ObjectSpaceId != null)
            .GroupBy(revision => revision.ObjectSpaceId!)
            .ToDictionary(group => group.Key, group => group.ToList());
        _objectsByRevision = store.Objects.Where(item => item.RevisionId != null)
            .GroupBy(item => item.RevisionId!)
            .ToDictionary(group => group.Key, group => group.ToList());
        _fileData = store.FileDataObjects.GroupBy(item => item.Id).ToDictionary(group => group.Key, group => group.Last());
    }

    public OneNoteMaterializedObjectSpace FindCurrentSpaceByRootJcid(uint jcid, string errorCode, string errorMessage) {
        foreach (OneNoteExtendedGuid id in _spaces.Keys) {
            if (id.Identifier == HeaderCellObjectSpaceId && id.Value == 1) continue;
            OneNoteMaterializedObjectSpace? space = TryGetCurrentSpace(id);
            if (space?.GetRoot(1)?.Jcid.Value == jcid) return space;
        }
        throw new OneNoteFormatException(errorCode, errorMessage);
    }

    public OneNoteMaterializedObjectSpace? TryGetCurrentSpace(OneNoteExtendedGuid id) => TryGetSpace(id, null);

    internal IReadOnlyCollection<OneNoteExtendedGuid> MappedObjectIds => _mappedObjectIds;

    public OneNoteMaterializedObjectSpace? TryGetSpace(OneNoteExtendedGuid id, OneNoteExtendedGuid? contextId) {
        string key = GetSpaceKey(id, contextId);
        if (_cache.TryGetValue(key, out OneNoteMaterializedObjectSpace? cached)) return cached;
        if (!_spaces.TryGetValue(id, out List<OneNoteRevisionManifest>? revisions)) return null;
        OneNoteRevisionManifest? current = revisions
            .SelectMany(revision => revision.RoleAssociations.Select(association => new { Revision = revision, Association = association }))
            .Where(item => item.Association.Role == 1 && ContextEquals(item.Association.ContextId, contextId) && !item.Revision.IsEncrypted)
            .OrderBy(item => item.Association.SourceOrder)
            .Select(item => item.Revision)
            .LastOrDefault();
        if (current == null) return null;

        var objects = new Dictionary<OneNoteExtendedGuid, OneNoteRevisionStoreObject>();
        var roots = new Dictionary<uint, OneNoteExtendedGuid>();
        foreach (OneNoteRevisionManifest revision in GetRevisionChain(current)) {
            if (_objectsByRevision.TryGetValue(revision.Id, out List<OneNoteRevisionStoreObject>? declarations)) {
                foreach (OneNoteRevisionStoreObject declaration in declarations) objects[declaration.Id] = declaration;
            }
            foreach (OneNoteRootObjectReference root in revision.RootObjects) roots[root.Role] = root.ObjectId;
        }
        var result = new OneNoteMaterializedObjectSpace(current, objects, roots, id => _mappedObjectIds.Add(id));
        _cache[key] = result;
        return result;
    }

    public static string GetSpaceKey(OneNoteExtendedGuid id, OneNoteExtendedGuid? contextId) {
        return id.ToString() + "|" + (contextId?.ToString() ?? "default");
    }

    public IReadOnlyList<OneNoteRevisionManifest> GetRevisionChain(OneNoteRevisionManifest revision) {
        var reversed = new List<OneNoteRevisionManifest>();
        var visited = new HashSet<OneNoteExtendedGuid>();
        OneNoteRevisionManifest? current = revision;
        while (current != null && visited.Add(current.Id)) {
            reversed.Add(current);
            current = current.DependencyId != null && _revisions.TryGetValue(current.DependencyId, out OneNoteRevisionManifest? dependency)
                ? dependency
                : null;
        }
        reversed.Reverse();
        return reversed.AsReadOnly();
    }

    public OneNoteBinaryPayload? ResolveFileData(OneNoteRevisionStoreObject item) {
        return TryResolveFileData(item, out _, out OneNoteBinaryPayload? payload) ? payload : null;
    }

    internal bool TryResolveFileData(OneNoteRevisionStoreObject item, out Guid id, out OneNoteBinaryPayload? payload) {
        id = Guid.Empty;
        payload = null;
        string? reference = item.FileDataReference;
        if (reference == null || !reference.StartsWith("<ifndf>", StringComparison.OrdinalIgnoreCase)) return false;
        string value = reference.Substring(7).Trim().TrimEnd('\0');
        if (!Guid.TryParse(value, out id) || !_fileData.TryGetValue(id, out OneNoteFileDataStoreObject? fileData)) return false;
        payload = fileData.Payload;
        return true;
    }

    private static bool ContextEquals(OneNoteExtendedGuid? left, OneNoteExtendedGuid? right) {
        return left == null ? right == null : left.Equals(right);
    }
}

internal sealed class OneNoteMaterializedObjectSpace {
    private readonly Dictionary<OneNoteExtendedGuid, OneNoteRevisionStoreObject> _objects;
    private readonly Dictionary<uint, OneNoteExtendedGuid> _roots;
    private readonly Action<OneNoteExtendedGuid> _markMapped;

    public OneNoteMaterializedObjectSpace(
        OneNoteRevisionManifest revision,
        Dictionary<OneNoteExtendedGuid, OneNoteRevisionStoreObject> objects,
        Dictionary<uint, OneNoteExtendedGuid> roots,
        Action<OneNoteExtendedGuid> markMapped) {
        Revision = revision;
        _objects = objects;
        _roots = roots;
        _markMapped = markMapped;
    }

    public OneNoteRevisionManifest Revision { get; }
    public IEnumerable<OneNoteRevisionStoreObject> Objects => _objects.Values;
    internal IReadOnlyDictionary<uint, OneNoteExtendedGuid> Roots => _roots;

    public OneNoteRevisionStoreObject? GetObject(OneNoteExtendedGuid id) {
        _objects.TryGetValue(id, out OneNoteRevisionStoreObject? value);
        if (value != null && OneNoteSemanticMapper.IsKnownJcid(value.Jcid)) _markMapped(id);
        return value;
    }

    public OneNoteRevisionStoreObject? GetRoot(uint role) {
        return _roots.TryGetValue(role, out OneNoteExtendedGuid? id) ? GetObject(id) : null;
    }
}
