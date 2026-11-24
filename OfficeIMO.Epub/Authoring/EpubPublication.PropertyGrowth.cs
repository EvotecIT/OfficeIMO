namespace OfficeIMO.Epub;

public sealed partial class EpubPublication {
    // Write derives svg, mathml and remote-resources properties from content, so the package it writes
    // can be larger than the live one. Each declaration whose content an edit validated keeps what Write
    // will derive for it, and budget checks subtract the net growth from the limit. Write serializes the
    // live package before deriving properties, so net shrinkage earns no credit. A rejected edit drops
    // what it staged. A declaration whose properties change is re-measured, and one whose payload is
    // replaced by anything other than the validated bytes loses its reservation.
    private readonly Dictionary<XElement, Reservation> _propertyGrowth = new();
    private readonly Dictionary<string, List<XElement>> _propertyGrowthByPath = new(StringComparer.Ordinal);
    private readonly HashSet<XElement> _propertyGrowthChanged = new();
    private readonly HashSet<string> _propertyGrowthChangedPaths = new(StringComparer.Ordinal);
    private Dictionary<XElement, (ContentFlags? Flags, byte[]? Payload)>? _stagedPropertyGrowth;
    private long _propertyGrowthTotal;

    private readonly struct ContentFlags {
        internal ContentFlags(bool svg, bool mathMl, bool remote) { Svg = svg; MathMl = mathMl; Remote = remote; }
        internal bool Svg { get; }
        internal bool MathMl { get; }
        internal bool Remote { get; }
    }

    private readonly struct Reservation {
        internal Reservation(ContentFlags flags, long growth, string? path, byte[] payload) { Flags = flags; Growth = growth; Path = path; Payload = payload; }
        internal ContentFlags Flags { get; }
        internal long Growth { get; }
        internal string? Path { get; }
        internal byte[] Payload { get; }
    }

    // Null flags drop the reservation, for a declaration whose content Write won't analyze again.
    private void StagePropertyGrowth(XElement declaration, ContentFlags? flags, byte[]? payload) =>
        (_stagedPropertyGrowth ??= new Dictionary<XElement, (ContentFlags?, byte[]?)>())[declaration] = (flags, payload);

    // Net bytes Write adds to the live package for every validated declaration, including this edit's.
    private long PropertyGrowthAllowance() {
        RefreshPropertyGrowth();
        long total = _propertyGrowthTotal;
        if (_stagedPropertyGrowth != null) {
            foreach (var staged in _stagedPropertyGrowth) {
                if (_propertyGrowth.TryGetValue(staged.Key, out Reservation committed)) total -= committed.Growth;
                if (staged.Value.Flags is ContentFlags flags) total += PropertyGrowth(staged.Key, flags);
            }
        }
        return Math.Max(0, total);
    }

    private void CommitPropertyGrowth() {
        if (_stagedPropertyGrowth == null) return;
        foreach (var staged in _stagedPropertyGrowth) {
            if (staged.Value.Flags is ContentFlags flags && staged.Value.Payload != null) SetPropertyGrowth(staged.Key, flags, staged.Value.Payload);
            else RemovePropertyGrowth(staged.Key);
        }
        _stagedPropertyGrowth = null;
    }

    private void SetPropertyGrowth(XElement declaration, ContentFlags flags, byte[] payload) {
        RemovePropertyGrowth(declaration);
        long growth = PropertyGrowth(declaration, flags);
        string? path = new EpubManifestItem(declaration, PackagePath).Reference.ContainerPath;
        _propertyGrowth[declaration] = new Reservation(flags, growth, path, payload);
        _propertyGrowthTotal += growth;
        if (path == null) return;
        if (!_propertyGrowthByPath.TryGetValue(path, out List<XElement>? declarations)) _propertyGrowthByPath.Add(path, declarations = new List<XElement>(1));
        declarations.Add(declaration);
    }

    private void RemovePropertyGrowth(XElement declaration) {
        if (!_propertyGrowth.TryGetValue(declaration, out Reservation previous)) return;
        _propertyGrowthTotal -= previous.Growth;
        _propertyGrowth.Remove(declaration);
        if (previous.Path != null && _propertyGrowthByPath.TryGetValue(previous.Path, out List<XElement>? declarations)) {
            declarations.Remove(declaration);
            if (declarations.Count == 0) _propertyGrowthByPath.Remove(previous.Path);
        }
    }

    // Called by the entry table whenever a retained payload is added, replaced or removed.
    private void OnEntryChanged(string path) {
        if (_propertyGrowthByPath.ContainsKey(path)) _propertyGrowthChangedPaths.Add(path);
    }

    private void RefreshPropertyGrowth() {
        if (_propertyGrowthChangedPaths.Count != 0) {
            foreach (string path in _propertyGrowthChangedPaths) {
                if (!_propertyGrowthByPath.TryGetValue(path, out List<XElement>? declarations)) continue;
                _entries.TryGetValue(path, out byte[] current);
                foreach (XElement declaration in declarations.ToArray())
                    if (!ReferenceEquals(_propertyGrowth[declaration].Payload, current)) RemovePropertyGrowth(declaration);
            }
            _propertyGrowthChangedPaths.Clear();
        }
        if (_propertyGrowthChanged.Count == 0) return;
        foreach (XElement declaration in _propertyGrowthChanged) {
            if (!_propertyGrowth.TryGetValue(declaration, out Reservation reserved)) continue;
            if (declaration.Document == _package && declaration.Parent?.Name == Opf + "manifest" && declaration.Name == Opf + "item" &&
                new EpubManifestItem(declaration, PackagePath).Reference.ContainerPath == reserved.Path) {
                SetPropertyGrowth(declaration, reserved.Flags, reserved.Payload);
            } else {
                RemovePropertyGrowth(declaration);
            }
        }
        _propertyGrowthChanged.Clear();
    }

    private long PropertyGrowth(XElement declaration, ContentFlags flags) {
        var item = new EpubManifestItem(declaration, PackagePath);
        return PropertiesBytes(ContentProperties(item, flags.Svg, flags.MathMl, flags.Remote)) - PropertiesBytes(item.Properties);
    }

    private static long PropertiesBytes(string? properties) =>
        properties == null ? 0 : Encoding.UTF8.GetByteCount(" properties=\"" + properties + "\"");

    // Called before a package change. Removing a section that holds declarations re-measures them all.
    private void MarkPropertyGrowthChanging(object? sender, XObjectChange change) {
        if (_propertyGrowth.Count == 0) return;
        switch (sender) {
            case XAttribute attribute when attribute.Parent != null && _propertyGrowth.ContainsKey(attribute.Parent):
                _propertyGrowthChanged.Add(attribute.Parent);
                break;
            case XElement element when change == XObjectChange.Remove || change == XObjectChange.Name:
                if (_propertyGrowth.ContainsKey(element)) _propertyGrowthChanged.Add(element);
                else if (element.HasElements) MarkAllPropertyGrowthChanged();
                break;
        }
    }

    // An attribute added to a declaration has its parent only once it's attached.
    private void MarkPropertyGrowthChanged(object? sender) {
        if (_propertyGrowth.Count != 0 && sender is XAttribute { Parent: XElement parent } && _propertyGrowth.ContainsKey(parent))
            _propertyGrowthChanged.Add(parent);
    }

    private void MarkAllPropertyGrowthChanged() {
        foreach (XElement declaration in _propertyGrowth.Keys) _propertyGrowthChanged.Add(declaration);
    }
}
