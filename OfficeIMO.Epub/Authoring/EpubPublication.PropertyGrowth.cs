namespace OfficeIMO.Epub;

public sealed partial class EpubPublication {
    // Write derives svg, mathml and remote-resources properties from content, so the package it writes
    // can be larger than the live one. Each declaration whose content an edit validated keeps what Write
    // will derive for it, and budget checks subtract the net growth from the limit. Write serializes the
    // live package before deriving properties, so net shrinkage earns no credit. A rejected edit drops
    // what it staged; a declaration whose properties change is re-measured.
    private readonly Dictionary<XElement, (ContentFlags Flags, long Growth)> _propertyGrowth = new();
    private readonly HashSet<XElement> _propertyGrowthChanged = new();
    private Dictionary<XElement, ContentFlags?>? _stagedPropertyGrowth;
    private long _propertyGrowthTotal;

    private readonly struct ContentFlags {
        internal ContentFlags(bool svg, bool mathMl, bool remote) { Svg = svg; MathMl = mathMl; Remote = remote; }
        internal bool Svg { get; }
        internal bool MathMl { get; }
        internal bool Remote { get; }
    }

    // Null flags drop the reservation, for a declaration whose content Write won't analyze again.
    private void StagePropertyGrowth(XElement declaration, ContentFlags? flags) =>
        (_stagedPropertyGrowth ??= new Dictionary<XElement, ContentFlags?>())[declaration] = flags;

    // Net bytes Write adds to the live package for every validated declaration, including this edit's.
    private long PropertyGrowthAllowance() {
        RefreshPropertyGrowth();
        long total = _propertyGrowthTotal;
        if (_stagedPropertyGrowth != null) {
            foreach (var staged in _stagedPropertyGrowth) {
                if (_propertyGrowth.TryGetValue(staged.Key, out var committed)) total -= committed.Growth;
                if (staged.Value is ContentFlags flags) total += PropertyGrowth(staged.Key, flags);
            }
        }
        return Math.Max(0, total);
    }

    private void CommitPropertyGrowth() {
        if (_stagedPropertyGrowth == null) return;
        foreach (var staged in _stagedPropertyGrowth) {
            if (staged.Value is ContentFlags flags) SetPropertyGrowth(staged.Key, flags);
            else if (_propertyGrowth.TryGetValue(staged.Key, out var previous)) {
                _propertyGrowthTotal -= previous.Growth;
                _propertyGrowth.Remove(staged.Key);
            }
        }
        _stagedPropertyGrowth = null;
    }

    private void SetPropertyGrowth(XElement declaration, ContentFlags flags) {
        if (_propertyGrowth.TryGetValue(declaration, out var previous)) _propertyGrowthTotal -= previous.Growth;
        long growth = PropertyGrowth(declaration, flags);
        _propertyGrowth[declaration] = (flags, growth);
        _propertyGrowthTotal += growth;
    }

    private void RefreshPropertyGrowth() {
        if (_propertyGrowthChanged.Count == 0) return;
        foreach (XElement declaration in _propertyGrowthChanged) {
            if (!_propertyGrowth.TryGetValue(declaration, out var reserved)) continue;
            if (declaration.Document == _package && declaration.Parent?.Name == Opf + "manifest" && declaration.Name == Opf + "item") {
                SetPropertyGrowth(declaration, reserved.Flags);
            } else {
                _propertyGrowthTotal -= reserved.Growth;
                _propertyGrowth.Remove(declaration);
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
