namespace OfficeIMO.AsciiDoc;

/// <summary>Kind of entry found in an AsciiDoc element attribute list.</summary>
public enum AsciiDocElementAttributeKind {
    /// <summary>Unnamed positional value.</summary>
    Positional = 0,
    /// <summary>Named key/value entry.</summary>
    Named,
    /// <summary>Shorthand ID beginning with <c>#</c>.</summary>
    Id,
    /// <summary>Shorthand role beginning with <c>.</c>.</summary>
    Role,
    /// <summary>Shorthand option beginning with <c>%</c>.</summary>
    Option
}

/// <summary>Parsed entry in an element attribute list.</summary>
public sealed class AsciiDocElementAttribute {
    internal AsciiDocElementAttribute(AsciiDocElementAttributeKind kind, int position, string rawText, string? name, string value) {
        Kind = kind;
        Position = position;
        RawText = rawText;
        Name = name;
        Value = value;
    }

    /// <summary>Entry kind.</summary>
    public AsciiDocElementAttributeKind Kind { get; }

    /// <summary>Zero-based position in the list.</summary>
    public int Position { get; }

    /// <summary>Exact entry text excluding surrounding comma separators.</summary>
    public string RawText { get; }

    /// <summary>Named entry key, or null for positional and shorthand entries.</summary>
    public string? Name { get; }

    /// <summary>Unquoted semantic value.</summary>
    public string Value { get; }
}

/// <summary>Semantic view of one AsciiDoc element attribute list.</summary>
public sealed class AsciiDocElementAttributes {
    private readonly IReadOnlyList<AsciiDocElementAttribute> _entries;

    internal AsciiDocElementAttributes(string source, IReadOnlyList<AsciiDocElementAttribute> entries) {
        Source = source;
        _entries = entries;
    }

    /// <summary>Raw content between the square brackets.</summary>
    public string Source { get; }

    /// <summary>Entries in source order.</summary>
    public IReadOnlyList<AsciiDocElementAttribute> Entries => _entries;

    /// <summary>First positional value, commonly the block style.</summary>
    public string? Style => Entries.FirstOrDefault(static entry => entry.Kind == AsciiDocElementAttributeKind.Positional)?.Value;

    /// <summary>Effective ID from shorthand or a named <c>id</c> entry.</summary>
    public string? Id => Entries.LastOrDefault(static entry => entry.Kind == AsciiDocElementAttributeKind.Id)?.Value
        ?? GetNamedValue("id");

    /// <summary>Roles declared using shorthand or a named <c>role</c> entry.</summary>
    public IReadOnlyList<string> Roles => CollectValues(AsciiDocElementAttributeKind.Role, "role");

    /// <summary>Options declared using shorthand or a named <c>options</c> entry.</summary>
    public IReadOnlyList<string> Options => CollectValues(AsciiDocElementAttributeKind.Option, "options");

    /// <summary>Gets the last named value using case-insensitive lookup.</summary>
    public string? GetNamedValue(string name) {
        if (name == null) throw new ArgumentNullException(nameof(name));
        for (int index = Entries.Count - 1; index >= 0; index--) {
            AsciiDocElementAttribute entry = Entries[index];
            if (entry.Kind == AsciiDocElementAttributeKind.Named && string.Equals(entry.Name, name, StringComparison.OrdinalIgnoreCase)) {
                return entry.Value;
            }
        }
        return null;
    }

    private IReadOnlyList<string> CollectValues(AsciiDocElementAttributeKind shorthandKind, string namedKey) {
        var values = new List<string>();
        for (int index = 0; index < Entries.Count; index++) {
            AsciiDocElementAttribute entry = Entries[index];
            if (entry.Kind == shorthandKind) {
                if (entry.Value.Length > 0) values.Add(entry.Value);
            } else if (entry.Kind == AsciiDocElementAttributeKind.Named &&
                       string.Equals(entry.Name, namedKey, StringComparison.OrdinalIgnoreCase)) {
                string[] parts = entry.Value.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                values.AddRange(parts);
            }
        }
        return values;
    }
}
