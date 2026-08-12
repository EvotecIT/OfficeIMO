namespace OfficeIMO.Html;

internal sealed class CascadeLayerRegistry {
    private readonly Dictionary<string, Dictionary<string, int>> _children = new Dictionary<string, Dictionary<string, int>>(StringComparer.Ordinal);
    private int _anonymous;

    internal CascadeLayerOrder Register(string name) {
        string[] components = name.Split(new[] { '.' }, StringSplitOptions.RemoveEmptyEntries);
        var positions = new List<int>(components.Length);
        string parent = string.Empty;
        foreach (string component in components) {
            if (!_children.TryGetValue(parent, out Dictionary<string, int>? children)) {
                children = new Dictionary<string, int>(StringComparer.Ordinal);
                _children[parent] = children;
            }
            if (!children.TryGetValue(component, out int position)) {
                position = children.Count;
                children[component] = position;
            }
            positions.Add(position);
            parent = parent.Length == 0 ? component : parent + "." + component;
        }
        return new CascadeLayerOrder(positions);
    }

    internal string RegisterAnonymous(string? parent) {
        string name = (parent == null ? string.Empty : parent + ".") + "#anonymous-" + (++_anonymous).ToString(System.Globalization.CultureInfo.InvariantCulture);
        Register(name);
        return name;
    }

    internal void RegisterStatement(string? prelude, string? parent) {
        foreach (string name in (prelude ?? string.Empty).Split(',')) {
            string trimmed = name.Trim();
            if (trimmed.Length > 0) Register(Combine(parent, trimmed));
        }
    }

    internal (string Path, CascadeLayerOrder Order) RegisterBlock(string? prelude, string? parent) {
        string name = (prelude ?? string.Empty).Trim();
        string path = name.Length == 0 ? RegisterAnonymous(parent) : Combine(parent, name);
        return (path, GetOrder(path));
    }

    internal CascadeLayerOrder GetOrder(string name) => Register(name);

    private static string Combine(string? parent, string child) =>
        string.IsNullOrWhiteSpace(parent) ? child : parent + "." + child;
}

internal sealed class CascadeLayerOrder : IEquatable<CascadeLayerOrder> {
    private readonly int[] _components;

    internal CascadeLayerOrder(IEnumerable<int> components) {
        _components = components.ToArray();
    }

    internal int CompareTo(CascadeLayerOrder other) {
        int shared = Math.Min(_components.Length, other._components.Length);
        for (int index = 0; index < shared; index++) {
            if (_components[index] != other._components[index]) return _components[index].CompareTo(other._components[index]);
        }

        // Declarations directly in a layer have normal precedence over declarations
        // in its nested sublayers. Important declarations reverse this ordering later.
        return other._components.Length.CompareTo(_components.Length);
    }

    public bool Equals(CascadeLayerOrder? other) =>
        other != null && _components.SequenceEqual(other._components);

    public override bool Equals(object? obj) => Equals(obj as CascadeLayerOrder);

    public override int GetHashCode() {
        unchecked {
            int hash = 17;
            foreach (int component in _components) hash = (hash * 31) + component;
            return hash;
        }
    }
}
