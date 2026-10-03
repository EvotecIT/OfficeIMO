namespace OfficeIMO.AsciiDoc;

/// <summary>Effective case-insensitive document attribute values.</summary>
public sealed class AsciiDocDocumentAttributes {
    private readonly Dictionary<string, string> _values;

    internal AsciiDocDocumentAttributes(Dictionary<string, string> values) {
        _values = new Dictionary<string, string>(values, StringComparer.OrdinalIgnoreCase);
    }

    /// <summary>Creates an attribute snapshot from caller-supplied values.</summary>
    public static AsciiDocDocumentAttributes Create(IReadOnlyDictionary<string, string>? values = null) {
        var copy = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (values != null) foreach (KeyValuePair<string, string> value in values) copy[value.Key] = value.Value;
        return new AsciiDocDocumentAttributes(copy);
    }

    /// <summary>Applies an assignment or unset, optionally resolving value references against the preceding snapshot.</summary>
    public AsciiDocDocumentAttributes Apply(AsciiDocAttributeEntry entry, bool expandReferences = false) {
        if (entry == null) throw new ArgumentNullException(nameof(entry));
        Dictionary<string, string> values = ToMutableDictionary();
        if (entry.IsUnset) values.Remove(entry.Name);
        else values[entry.Name] = expandReferences ? AsciiDocAttributeSubstitutor.Substitute(entry.Value, this).Value : entry.Value;
        return new AsciiDocDocumentAttributes(values);
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
