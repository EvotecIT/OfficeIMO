namespace OfficeIMO.Markdown;

/// <summary>
/// Generic Markdown attributes attached to semantic objects and syntax nodes.
/// </summary>
public sealed class MarkdownAttributeSet {
    private static readonly IReadOnlyList<string> EmptyClasses = Array.Empty<string>();
    private static readonly IReadOnlyDictionary<string, string?> EmptyAttributeMap =
        new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);

    private readonly IReadOnlyList<string> _classes;
    private readonly IReadOnlyDictionary<string, string?> _attributes;

    private MarkdownAttributeSet(
        string? elementId,
        IReadOnlyList<string> classes,
        IReadOnlyDictionary<string, string?> attributes) {
        ElementId = string.IsNullOrWhiteSpace(elementId) ? null : elementId!.Trim();
        _classes = classes;
        _attributes = attributes;
    }

    /// <summary>Empty attribute set used by nodes without generic attributes.</summary>
    public static MarkdownAttributeSet Empty { get; } =
        new MarkdownAttributeSet(null, EmptyClasses, EmptyAttributeMap);

    /// <summary>Optional element id associated with the Markdown node.</summary>
    public string? ElementId { get; }

    /// <summary>CSS-like classes associated with the Markdown node.</summary>
    public IReadOnlyList<string> Classes => _classes;

    /// <summary>Additional key/value attributes associated with the Markdown node.</summary>
    public IReadOnlyDictionary<string, string?> Attributes => _attributes;

    /// <summary>Whether the set has no id, classes, or key/value attributes.</summary>
    public bool IsEmpty => ElementId == null && _classes.Count == 0 && _attributes.Count == 0;

    /// <summary>Create a generic Markdown attribute set.</summary>
    public static MarkdownAttributeSet Create(
        string? elementId = null,
        IEnumerable<string>? classes = null,
        IEnumerable<KeyValuePair<string, string?>>? attributes = null) {
        var classList = CopyClasses(classes);
        var attributeMap = CopyAttributes(attributes);

        if (string.IsNullOrWhiteSpace(elementId) && classList.Count == 0 && attributeMap.Count == 0) {
            return Empty;
        }

        return new MarkdownAttributeSet(elementId, classList, attributeMap);
    }

    /// <summary>Reads an attribute value by key.</summary>
    public bool TryGetAttribute(string name, out string? value) {
        value = null;
        if (string.IsNullOrWhiteSpace(name) || _attributes.Count == 0) {
            return false;
        }

        return _attributes.TryGetValue(name.Trim(), out value);
    }

    /// <summary>Reads the first matching attribute value by alias.</summary>
    public string? GetAttribute(params string[] aliases) {
        if (aliases == null || aliases.Length == 0 || _attributes.Count == 0) {
            return null;
        }

        for (int i = 0; i < aliases.Length; i++) {
            if (TryGetAttribute(aliases[i], out var value)) {
                return value;
            }
        }

        return null;
    }

    /// <summary>Checks whether a class is present using ordinal comparison.</summary>
    public bool HasClass(string className) {
        if (string.IsNullOrWhiteSpace(className) || _classes.Count == 0) {
            return false;
        }

        for (int i = 0; i < _classes.Count; i++) {
            if (string.Equals(_classes[i], className.Trim(), StringComparison.Ordinal)) {
                return true;
            }
        }

        return false;
    }

    private static IReadOnlyList<string> CopyClasses(IEnumerable<string>? classes) {
        if (classes == null) {
            return EmptyClasses;
        }

        var result = new List<string>();
        foreach (var className in classes) {
            if (string.IsNullOrWhiteSpace(className)) {
                continue;
            }

            result.Add(className.Trim());
        }

        return result.Count == 0 ? EmptyClasses : result;
    }

    private static IReadOnlyDictionary<string, string?> CopyAttributes(IEnumerable<KeyValuePair<string, string?>>? attributes) {
        if (attributes == null) {
            return EmptyAttributeMap;
        }

        var result = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);
        foreach (var attribute in attributes) {
            if (string.IsNullOrWhiteSpace(attribute.Key)) {
                continue;
            }

            result[attribute.Key.Trim()] = attribute.Value;
        }

        return result.Count == 0 ? EmptyAttributeMap : result;
    }
}
