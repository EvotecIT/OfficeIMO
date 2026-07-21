using System.Text.Json;

namespace OfficeIMO.Adf;

/// <summary>
/// Represents an Atlas Document Format document while retaining fields that OfficeIMO does not yet understand.
/// </summary>
public sealed class AdfDocument {
    /// <summary>Creates an empty ADF document.</summary>
    public AdfDocument() {
    }

    /// <summary>Creates an ADF document with the supplied top-level nodes.</summary>
    public AdfDocument(IEnumerable<AdfNode>? content) {
        if (content != null) {
            Content.AddRange(content);
        }
    }

    /// <summary>ADF schema version.</summary>
    public int Version { get; set; } = 1;

    /// <summary>Root node type. A normal ADF document uses <c>doc</c>.</summary>
    public string Type { get; set; } = "doc";

    /// <summary>Top-level document nodes.</summary>
    public List<AdfNode> Content { get; } = new List<AdfNode>();

    /// <summary>Unknown root properties retained during parse/write round trips.</summary>
    public IDictionary<string, JsonElement> ExtensionData { get; } = new Dictionary<string, JsonElement>(StringComparer.Ordinal);

    /// <summary>Parses an ADF JSON document.</summary>
    public static AdfDocument Parse(string json) => AdfJsonSerializer.Parse(json);

    /// <summary>Serializes this document to ADF JSON.</summary>
    public string ToJson(bool indented = false) => AdfJsonSerializer.Serialize(this, indented);

    /// <summary>Validates the structural ADF contract without rejecting unknown node or mark types.</summary>
    public AdfValidationResult Validate() => AdfValidator.Validate(this);
}

/// <summary>An ADF content node.</summary>
public sealed class AdfNode {
    /// <summary>Creates an ADF node.</summary>
    public AdfNode(string type) {
        if (string.IsNullOrWhiteSpace(type)) throw new ArgumentException("ADF node type is required.", nameof(type));
        Type = type;
    }

    /// <summary>Node type such as <c>paragraph</c>, <c>heading</c>, or <c>text</c>.</summary>
    public string Type { get; set; }

    /// <summary>Text payload for text nodes.</summary>
    public string? Text { get; set; }

    /// <summary>Child nodes.</summary>
    public List<AdfNode> Content { get; } = new List<AdfNode>();

    /// <summary>Marks applied to a text node.</summary>
    public List<AdfMark> Marks { get; } = new List<AdfMark>();

    /// <summary>Node attributes retained as arbitrary JSON values.</summary>
    public IDictionary<string, JsonElement> Attributes { get; } = new Dictionary<string, JsonElement>(StringComparer.Ordinal);

    /// <summary>Unknown node properties retained during parse/write round trips.</summary>
    public IDictionary<string, JsonElement> ExtensionData { get; } = new Dictionary<string, JsonElement>(StringComparer.Ordinal);

    /// <summary>Creates a text node.</summary>
    public static AdfNode TextNode(string text, IEnumerable<AdfMark>? marks = null) {
        var node = new AdfNode("text") { Text = text ?? string.Empty };
        if (marks != null) node.Marks.AddRange(marks);
        return node;
    }

    /// <summary>Sets a JSON-compatible attribute value.</summary>
    public AdfNode SetAttribute<T>(string name, T value) {
        if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Attribute name is required.", nameof(name));
        Attributes[name] = JsonSerializer.SerializeToElement(value).Clone();
        return this;
    }

    /// <summary>Gets a string attribute when present.</summary>
    public string? GetStringAttribute(string name) =>
        Attributes.TryGetValue(name, out JsonElement value) && value.ValueKind == JsonValueKind.String
            ? value.GetString()
            : null;

    /// <summary>Gets an integer attribute when present.</summary>
    public int? GetInt32Attribute(string name) =>
        Attributes.TryGetValue(name, out JsonElement value) && value.ValueKind == JsonValueKind.Number && value.TryGetInt32(out int result)
            ? result
            : null;
}

/// <summary>An ADF text mark such as strong, emphasis, code, or link.</summary>
public sealed class AdfMark {
    /// <summary>Creates an ADF mark.</summary>
    public AdfMark(string type) {
        if (string.IsNullOrWhiteSpace(type)) throw new ArgumentException("ADF mark type is required.", nameof(type));
        Type = type;
    }

    /// <summary>Mark type.</summary>
    public string Type { get; set; }

    /// <summary>Mark attributes retained as arbitrary JSON values.</summary>
    public IDictionary<string, JsonElement> Attributes { get; } = new Dictionary<string, JsonElement>(StringComparer.Ordinal);

    /// <summary>Unknown mark properties retained during parse/write round trips.</summary>
    public IDictionary<string, JsonElement> ExtensionData { get; } = new Dictionary<string, JsonElement>(StringComparer.Ordinal);

    /// <summary>Sets a JSON-compatible mark attribute.</summary>
    public AdfMark SetAttribute<T>(string name, T value) {
        if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Attribute name is required.", nameof(name));
        Attributes[name] = JsonSerializer.SerializeToElement(value).Clone();
        return this;
    }

    /// <summary>Gets a string attribute when present.</summary>
    public string? GetStringAttribute(string name) =>
        Attributes.TryGetValue(name, out JsonElement value) && value.ValueKind == JsonValueKind.String
            ? value.GetString()
            : null;
}
