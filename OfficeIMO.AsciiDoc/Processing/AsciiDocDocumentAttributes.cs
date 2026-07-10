namespace OfficeIMO.AsciiDoc;

/// <summary>Effective case-insensitive document attribute values.</summary>
public sealed class AsciiDocDocumentAttributes {
    private readonly Dictionary<string, string> _values;

    internal AsciiDocDocumentAttributes(Dictionary<string, string> values) {
        _values = new Dictionary<string, string>(values, StringComparer.OrdinalIgnoreCase);
    }

    /// <summary>Number of set attributes.</summary>
    public int Count => _values.Count;

    /// <summary>Set attribute names and values.</summary>
    public IReadOnlyDictionary<string, string> Values => _values;

    /// <summary>Tests whether an attribute is set.</summary>
    public bool Contains(string name) {
        if (name == null) throw new ArgumentNullException(nameof(name));
        return _values.ContainsKey(name);
    }

    /// <summary>Gets an attribute value when set.</summary>
    public bool TryGetValue(string name, out string value) {
        if (name == null) throw new ArgumentNullException(nameof(name));
        return _values.TryGetValue(name, out value!);
    }

    /// <summary>Gets an attribute value or null.</summary>
    public string? GetValueOrDefault(string name) => TryGetValue(name, out string value) ? value : null;

    internal Dictionary<string, string> ToMutableDictionary() =>
        new Dictionary<string, string>(_values, StringComparer.OrdinalIgnoreCase);
}
