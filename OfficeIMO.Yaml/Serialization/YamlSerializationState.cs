using System.Runtime.CompilerServices;

namespace OfficeIMO.Yaml;

/// <summary>Compares objects by reference, for anchors and cycle detection.</summary>
internal sealed class YamlReferenceComparer : IEqualityComparer<object> {
    public static YamlReferenceComparer Instance { get; } = new YamlReferenceComparer();

    public new bool Equals(object? x, object? y) => ReferenceEquals(x, y);

    public int GetHashCode(object obj) => RuntimeHelpers.GetHashCode(obj);
}

/// <summary>Per-call state of one deserialization: anchored instances and the path for error messages.</summary>
internal sealed class YamlReadState {
    private struct PathStep {
        public string? Key;
        public int Index;
    }

    private PathStep[] _path = new PathStep[4];
    private int _depth;
    private Dictionary<string, object?>? _references;

    public YamlReadState(YamlSerializerOptions options) {
        Options = options;
        Preserve = options.ReferenceHandling == YamlReferenceHandling.Preserve;
    }

    public YamlSerializerOptions Options { get; }

    public bool Preserve { get; }

    /// <summary>The anchor of the node being read, until the converter that creates its instance takes it.</summary>
    public string? PendingAnchor { get; set; }

    public void PushKey(string key) {
        if (_depth == _path.Length) Array.Resize(ref _path, _depth * 2);
        _path[_depth++] = new PathStep { Key = key };
    }

    public void PushIndex(int index) {
        if (_depth == _path.Length) Array.Resize(ref _path, _depth * 2);
        _path[_depth++] = new PathStep { Index = index };
    }

    public void SetIndex(int index) => _path[_depth - 1].Index = index;

    public void Pop() => _depth--;

    /// <summary>Registers the instance created for the pending anchor, before its children are read, so cycles resolve.</summary>
    public void TakePendingAnchor(object instance) {
        if (PendingAnchor is null) return;
        (_references ??= new Dictionary<string, object?>(StringComparer.Ordinal))[PendingAnchor] = instance;
        PendingAnchor = null;
    }

    public void Register(string anchor, object? value) =>
        (_references ??= new Dictionary<string, object?>(StringComparer.Ordinal))[anchor] = value;

    public bool TryGetReference(string anchor, out object? value) {
        value = null;
        return _references is not null && _references.TryGetValue(anchor, out value);
    }

    /// <summary>Formats the path of the node being read, in the path grammar of <see cref="YamlNode.Select"/>.</summary>
    public string FormatPath() {
        if (_depth == 0) return string.Empty;
        var text = new StringBuilder();
        for (int i = 0; i < _depth; i++) {
            PathStep step = _path[i];
            if (step.Key is null) {
                text.Append('[').Append(step.Index.ToString(CultureInfo.InvariantCulture)).Append(']');
            } else if (IsSimpleName(step.Key)) {
                if (text.Length > 0) text.Append('.');
                text.Append(step.Key);
            } else {
                text.Append("[\"").Append(step.Key.Replace("\\", "\\\\").Replace("\"", "\\\"")).Append("\"]");
            }
        }
        return text.ToString();
    }

    private static bool IsSimpleName(string key) {
        if (key.Length == 0) return false;
        foreach (char c in key) {
            if (c is '.' or '[' or ']' or '"' or '$' || char.IsWhiteSpace(c) || char.IsControl(c)) return false;
        }
        return true;
    }
}

/// <summary>Per-call state of one serialization: anchors for shared references and the stack for cycle detection.</summary>
internal sealed class YamlWriteState {
    private object[] _ancestors = new object[16];
    private int _depth;
    private Dictionary<object, string>? _anchors;
    private HashSet<object>? _shared;
    private int _nextAnchor;

    public YamlWriteState(YamlSerializerOptions options) {
        Options = options;
        Preserve = options.ReferenceHandling == YamlReferenceHandling.Preserve;
    }

    public YamlSerializerOptions Options { get; }

    public bool Preserve { get; }

    /// <summary>The discriminator key and value the next object mapping writes first, for property discriminators.</summary>
    public string? DiscriminatorKey { get; set; }

    public string? DiscriminatorValue { get; set; }

    /// <summary>Sets the objects reached more than once, found by walking the graph before writing.</summary>
    public void SetShared(HashSet<object> shared) => _shared = shared;

    /// <summary>
    /// Returns the alias to write for an object already written, or null; for a shared object reached the first time it
    /// assigns the anchor to write.
    /// </summary>
    public bool TryGetAlias(object value, out string? alias, out string? anchor) {
        alias = null;
        anchor = null;
        if (_shared is null || !_shared.Contains(value)) return false;
        _anchors ??= new Dictionary<object, string>(YamlReferenceComparer.Instance);
        if (_anchors.TryGetValue(value, out alias)) return true;
        anchor = "id" + (++_nextAnchor).ToString("000", CultureInfo.InvariantCulture);
        _anchors.Add(value, anchor);
        return false;
    }

    /// <summary>
    /// Starts writing an object or collection. Under Preserve it writes the alias of an object already written and returns
    /// false, or sets the anchor of a shared object; otherwise it fails on a cycle. Pair a true result with <see cref="End"/>.
    /// </summary>
    public bool Begin(YamlWriter writer, object value, Type type) {
        if (Preserve) {
            if (TryGetAlias(value, out string? alias, out string? anchor)) {
                writer.WriteAlias(alias!);
                return false;
            }
            if (anchor is not null) writer.WriteAnchor(anchor);
            return true;
        }
        Enter(value, type);
        return true;
    }

    public void End() {
        if (!Preserve) Exit();
    }

    /// <summary>Enters an object or collection, failing on a cycle.</summary>
    public void Enter(object value, Type type) {
        for (int i = 0; i < _depth; i++) {
            if (ReferenceEquals(_ancestors[i], value)) throw YamlSerializationErrors.Cycle(type);
        }
        if (_depth == _ancestors.Length) Array.Resize(ref _ancestors, _depth * 2);
        _ancestors[_depth++] = value;
    }

    public void Exit() => _ancestors[--_depth] = null!;
}

/// <summary>Walks an object graph before writing with <see cref="YamlReferenceHandling.Preserve"/> to find shared references.</summary>
internal sealed class YamlReferenceCollector {
    private readonly HashSet<object> _seen = new HashSet<object>(YamlReferenceComparer.Instance);

    public YamlReferenceCollector(YamlSerializerOptions options) {
        Options = options;
    }

    public YamlSerializerOptions Options { get; }

    public HashSet<object> Shared { get; } = new HashSet<object>(YamlReferenceComparer.Instance);

    /// <summary>Returns true the first time an object is visited; a second visit marks it shared.</summary>
    public bool Visit(object value) {
        if (_seen.Add(value)) return true;
        Shared.Add(value);
        return false;
    }
}
