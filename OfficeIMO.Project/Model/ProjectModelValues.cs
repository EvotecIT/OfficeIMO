namespace OfficeIMO.Project;

/// <summary>Sparse scalar inventory; absent/null values do not allocate dictionary entries.</summary>
internal sealed class ProjectModelValues : Dictionary<string, object?> {
    internal ProjectModelValues() : base(StringComparer.Ordinal) { }
    public new object? this[string key] {
        get => TryGetValue(key, out var value) ? value : null;
        set { if (value == null) Remove(key); else base[key] = value; }
    }
}
