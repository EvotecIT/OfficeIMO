namespace OfficeIMO.AsciiDoc;

/// <summary>Effective case-insensitive document attribute values.</summary>
public sealed class AsciiDocDocumentAttributes {
    private readonly AsciiDocAttributeMap _values;

    internal AsciiDocDocumentAttributes(Dictionary<string, string> values)
        : this(new AsciiDocAttributeMap(values)) { }

    private AsciiDocDocumentAttributes(AsciiDocAttributeMap values) => _values = values;

    /// <summary>Creates an attribute snapshot from caller-supplied values.</summary>
    public static AsciiDocDocumentAttributes Create(IReadOnlyDictionary<string, string>? values = null) =>
        new AsciiDocDocumentAttributes(new AsciiDocAttributeMap(values));

    /// <summary>Applies an assignment or unset, optionally resolving value references against the preceding snapshot.</summary>
    public AsciiDocDocumentAttributes Apply(AsciiDocAttributeEntry entry, bool expandReferences = false) {
        if (entry == null) throw new ArgumentNullException(nameof(entry));
        if (entry.IsUnset) return new AsciiDocDocumentAttributes(_values.Remove(entry.Name));
        string value = expandReferences ? AsciiDocAttributeSubstitutor.Substitute(entry.Value, this).Value : entry.Value;
        return new AsciiDocDocumentAttributes(_values.Set(entry.Name, value));
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

}
