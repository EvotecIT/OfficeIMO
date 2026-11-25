namespace OfficeIMO.Epub;

public sealed partial class EpubPublication {
    private ManifestLookup? _manifestLookup;
    private Dictionary<string, int>? _packageIds;
    private Dictionary<string, List<XElement>>? _spinePositions;

    // Reading positions by manifest id, kept current like the manifest lookup.
    private IReadOnlyList<XElement> SpinePositions(string manifestId) {
        if (_spinePositions == null) {
            _spinePositions = new Dictionary<string, List<XElement>>(StringComparer.Ordinal);
            foreach (XElement position in RequireSection("spine").Elements(Opf + "itemref")) AddSpinePosition(position);
        }
        return _spinePositions.TryGetValue(manifestId, out List<XElement>? positions) ? positions : (IReadOnlyList<XElement>)Array.Empty<XElement>();
    }

    private void AddSpinePosition(XElement position) {
        string? id = (string?)position.Attribute("idref");
        if (id == null || _spinePositions == null) return;
        if (!_spinePositions.TryGetValue(id, out List<XElement>? positions)) _spinePositions.Add(id, positions = new List<XElement>(1));
        positions.Add(position);
    }

    // Manifest declarations indexed by id and resolved path. The live lookup follows package changes:
    // appending a declaration extends it, and any other manifest change rebuilds it on next use.
    private ManifestLookup LiveManifest => _manifestLookup ??= new ManifestLookup(RequireSection("manifest").Elements(Opf + "item"), PackagePath);

    private bool PackageIdExists(string id) {
        if (_packageIds == null) {
            _packageIds = new Dictionary<string, int>(StringComparer.Ordinal);
            foreach (XAttribute attribute in Root.DescendantsAndSelf().Attributes("id")) CountPackageId(attribute.Value, 1);
        }
        return _packageIds.ContainsKey(id);
    }

    private void CountPackageId(string id, int change) {
        if (_packageIds == null) return;
        _packageIds.TryGetValue(id, out int count);
        count += change;
        if (count > 0) _packageIds[id] = count; else _packageIds.Remove(id);
    }

    private void CountPackageIds(XElement element, int change) {
        if (_packageIds == null) return;
        foreach (XAttribute attribute in element.DescendantsAndSelf().Attributes("id")) CountPackageId(attribute.Value, change);
    }

    private void InvalidatePackageIndexes() {
        _manifestLookup = null;
        _packageIds = null;
        _spinePositions = null;
        _emptyStringElements = null;
        MarkAllPropertyGrowthChanged();
    }

    private bool InManifest(XObject? node) => InSection(node, "manifest");

    private bool InSection(XObject? node, string section) {
        XElement? element = node as XElement ?? node?.Parent;
        for (; element != null; element = element.Parent)
            if (element.Name == Opf + section && element.Parent == Root) return true;
        return false;
    }

    // Called before a package change, while the changed object is still in its old state.
    private void IndexPackageChanging(object? sender, XObjectChange change) {
        switch (sender) {
            case XElement element when change == XObjectChange.Remove:
                CountPackageIds(element, -1);
                if (InManifest(element)) _manifestLookup = null;
                if (InSection(element, "spine")) _spinePositions = null;
                break;
            case XElement element when change == XObjectChange.Name:
                InvalidatePackageIndexes();
                break;
            case XAttribute attribute when change == XObjectChange.Remove || change == XObjectChange.Value:
                if (attribute.Name == "id") CountPackageId(attribute.Value, -1);
                if (InManifest(attribute)) _manifestLookup = null;
                if (attribute.Name == "idref" && InSection(attribute, "spine")) _spinePositions = null;
                break;
        }
    }

    // Called after a package change, once the changed object is in its new state.
    private void IndexPackageChanged(object? sender, XObjectChange change) {
        switch (sender) {
            case XElement element when change == XObjectChange.Add:
                CountPackageIds(element, 1);
                bool appended = element.NextNode == null && element.Parent?.Parent == Root;
                if (_manifestLookup != null) {
                    if (appended && element.Name == Opf + "item" && element.Parent!.Name == Opf + "manifest") _manifestLookup.Append(element);
                    else if (InManifest(element)) _manifestLookup = null;
                }
                if (_spinePositions != null) {
                    if (appended && element.Name == Opf + "itemref" && element.Parent!.Name == Opf + "spine") AddSpinePosition(element);
                    else if (InSection(element, "spine")) _spinePositions = null;
                }
                break;
            case XAttribute attribute when change == XObjectChange.Add || change == XObjectChange.Value:
                if (attribute.Name == "id") CountPackageId(attribute.Value, 1);
                if (InManifest(attribute)) _manifestLookup = null;
                if (attribute.Name == "idref" && InSection(attribute, "spine")) _spinePositions = null;
                break;
        }
    }

    private sealed class ManifestLookup {
        private readonly string _packagePath;
        private readonly List<EpubManifestItem> _items = new();
        private readonly Dictionary<string, EpubManifestItem> _byId = new(StringComparer.Ordinal);
        private readonly Dictionary<string, List<EpubManifestItem>> _byPath = new(StringComparer.Ordinal);
        private readonly Dictionary<string, EpubManifestItem> _remote = new(StringComparer.Ordinal);
        private readonly HashSet<string> _scripted = new(StringComparer.Ordinal);
        private readonly List<EpubManifestItem> _navigation = new();

        internal ManifestLookup(IEnumerable<XElement> items, string packagePath) {
            _packagePath = packagePath;
            foreach (XElement item in items) Append(item);
        }

        internal ManifestLookup(EpubManifestItem[] items, string packagePath) {
            _packagePath = packagePath;
            foreach (EpubManifestItem item in items) Append(item);
        }

        internal IReadOnlyList<EpubManifestItem> Items => _items;
        internal bool HasDuplicateIds { get; private set; }
        internal IReadOnlyList<EpubManifestItem> NavigationItems => _navigation;

        internal void Append(XElement element) => Append(new EpubManifestItem(element, _packagePath));

        private void Append(EpubManifestItem item) {
            _items.Add(item);
            if (_byId.ContainsKey(item.Id)) HasDuplicateIds = true; else _byId.Add(item.Id, item);
            if (HasToken(item.Properties, "nav")) _navigation.Add(item);
            EpubReference reference = item.Reference;
            if (reference.Kind == EpubReferenceKind.Container && reference.ContainerPath != null) {
                if (!_byPath.TryGetValue(reference.ContainerPath, out List<EpubManifestItem>? declarations))
                    _byPath.Add(reference.ContainerPath, declarations = new List<EpubManifestItem>(1));
                declarations.Add(item);
                if (HasToken(item.Properties, "scripted") || IsScriptMediaType(item.MediaType)) _scripted.Add(reference.ContainerPath);
            } else if (reference.Kind == EpubReferenceKind.External && reference.Target != null && !_remote.ContainsKey(reference.Target))
                _remote.Add(reference.Target, item);
        }

        internal EpubManifestItem? ById(string id) => _byId.TryGetValue(id, out EpubManifestItem? item) ? item : null;

        internal IReadOnlyList<EpubManifestItem> AtPath(string path) =>
            _byPath.TryGetValue(path, out List<EpubManifestItem>? items) ? items : (IReadOnlyList<EpubManifestItem>)Array.Empty<EpubManifestItem>();

        internal bool IsLocal(string path) => _byPath.ContainsKey(path);
        internal bool IsScripted(string path) => _scripted.Contains(path);
        internal EpubManifestItem? Remote(string target) => _remote.TryGetValue(target, out EpubManifestItem? item) ? item : null;
    }
}
