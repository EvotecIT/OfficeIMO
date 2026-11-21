namespace OfficeIMO.Yaml.Tests.Conformance;

/// <summary>Kinds of loaded values compared with the suite's JSON.</summary>
public enum YamlConformanceNodeKind {
    /// <summary>Null.</summary>
    Null,
    /// <summary>Boolean; <see cref="YamlConformanceNode.Scalar"/> is <c>true</c> or <c>false</c>.</summary>
    Boolean,
    /// <summary>Integer; <see cref="YamlConformanceNode.Scalar"/> is decimal digits with an optional leading minus.</summary>
    Integer,
    /// <summary>
    /// Floating point; <see cref="YamlConformanceNode.Scalar"/> parses with invariant culture
    /// (<c>Infinity</c>, <c>-Infinity</c> and <c>NaN</c> for the special values).
    /// </summary>
    Float,
    /// <summary>String; <see cref="YamlConformanceNode.Scalar"/> is the content.</summary>
    String,
    /// <summary>Sequence of <see cref="YamlConformanceNode.Items"/>.</summary>
    Sequence,
    /// <summary>Mapping of <see cref="YamlConformanceNode.Entries"/> in source order.</summary>
    Mapping
}

/// <summary>Engine-neutral loaded value: the shared shape of an engine's documents and the suite's <c>in.json</c>.</summary>
public sealed class YamlConformanceNode {
    private static readonly IReadOnlyList<YamlConformanceNode> NoItems = new YamlConformanceNode[0];
    private static readonly IReadOnlyList<KeyValuePair<YamlConformanceNode, YamlConformanceNode>> NoEntries = new KeyValuePair<YamlConformanceNode, YamlConformanceNode>[0];

    private YamlConformanceNode(YamlConformanceNodeKind kind, string? scalar, IReadOnlyList<YamlConformanceNode> items, IReadOnlyList<KeyValuePair<YamlConformanceNode, YamlConformanceNode>> entries) {
        Kind = kind;
        Scalar = scalar;
        Items = items;
        Entries = entries;
    }

    /// <summary>Value kind.</summary>
    public YamlConformanceNodeKind Kind { get; }

    /// <summary>Scalar text for null-free scalar kinds; null for null, sequences and mappings.</summary>
    public string? Scalar { get; }

    /// <summary>Sequence items; empty for other kinds.</summary>
    public IReadOnlyList<YamlConformanceNode> Items { get; }

    /// <summary>Mapping entries; empty for other kinds.</summary>
    public IReadOnlyList<KeyValuePair<YamlConformanceNode, YamlConformanceNode>> Entries { get; }

    /// <summary>The null value.</summary>
    public static YamlConformanceNode Null { get; } = new(YamlConformanceNodeKind.Null, null, NoItems, NoEntries);

    /// <summary>Creates a boolean.</summary>
    public static YamlConformanceNode Boolean(bool value) => new(YamlConformanceNodeKind.Boolean, value ? "true" : "false", NoItems, NoEntries);

    /// <summary>Creates an integer from decimal digits.</summary>
    public static YamlConformanceNode Integer(string digits) => new(YamlConformanceNodeKind.Integer, digits ?? throw new ArgumentNullException(nameof(digits)), NoItems, NoEntries);

    /// <summary>Creates a floating-point number from invariant text.</summary>
    public static YamlConformanceNode Float(string text) => new(YamlConformanceNodeKind.Float, text ?? throw new ArgumentNullException(nameof(text)), NoItems, NoEntries);

    /// <summary>Creates a string.</summary>
    public static YamlConformanceNode String(string value) => new(YamlConformanceNodeKind.String, value ?? throw new ArgumentNullException(nameof(value)), NoItems, NoEntries);

    /// <summary>Creates a sequence.</summary>
    public static YamlConformanceNode Sequence(IReadOnlyList<YamlConformanceNode> items) =>
        new(YamlConformanceNodeKind.Sequence, null, items ?? throw new ArgumentNullException(nameof(items)), NoEntries);

    /// <summary>Creates a mapping.</summary>
    public static YamlConformanceNode Mapping(IReadOnlyList<KeyValuePair<YamlConformanceNode, YamlConformanceNode>> entries) =>
        new(YamlConformanceNodeKind.Mapping, null, NoItems, entries ?? throw new ArgumentNullException(nameof(entries)));

    /// <summary>Returns the value of the first entry whose key is the given string, or null.</summary>
    public YamlConformanceNode? Get(string key) {
        foreach (KeyValuePair<YamlConformanceNode, YamlConformanceNode> entry in Entries) {
            if (entry.Key.Kind == YamlConformanceNodeKind.String && string.Equals(entry.Key.Scalar, key, StringComparison.Ordinal)) return entry.Value;
        }

        return null;
    }

    /// <summary>Returns the string value of a required entry.</summary>
    public string GetString(string key) {
        YamlConformanceNode? value = Get(key);
        if (value == null || value.Kind != YamlConformanceNodeKind.String) throw new FormatException("Expected a string property '" + key + "'.");
        return value.Scalar!;
    }

    /// <summary>Returns a short diagnostic form.</summary>
    public override string ToString() => Kind switch {
        YamlConformanceNodeKind.Null => "null",
        YamlConformanceNodeKind.String => "\"" + Scalar!.Replace("\\", "\\\\").Replace("\n", "\\n").Replace("\r", "\\r").Replace("\t", "\\t") + "\"",
        YamlConformanceNodeKind.Sequence => "[" + Items.Count + " items]",
        YamlConformanceNodeKind.Mapping => "{" + Entries.Count + " entries}",
        _ => Scalar ?? string.Empty
    };
}
