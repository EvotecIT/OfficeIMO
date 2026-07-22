using System.Collections;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.Json.Serialization.Metadata;

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
        Attributes[name] = AdfJsonValue.Create(value);
        return this;
    }

    /// <summary>Sets a custom attribute using source-generated JSON metadata.</summary>
    public AdfNode SetAttribute<T>(string name, T value, JsonTypeInfo<T> typeInfo) {
        if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Attribute name is required.", nameof(name));
        if (typeInfo == null) throw new ArgumentNullException(nameof(typeInfo));
        Attributes[name] = JsonSerializer.SerializeToElement(value, typeInfo).Clone();
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
        Attributes[name] = AdfJsonValue.Create(value);
        return this;
    }

    /// <summary>Sets a custom mark attribute using source-generated JSON metadata.</summary>
    public AdfMark SetAttribute<T>(string name, T value, JsonTypeInfo<T> typeInfo) {
        if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Attribute name is required.", nameof(name));
        if (typeInfo == null) throw new ArgumentNullException(nameof(typeInfo));
        Attributes[name] = JsonSerializer.SerializeToElement(value, typeInfo).Clone();
        return this;
    }

    /// <summary>Gets a string attribute when present.</summary>
    public string? GetStringAttribute(string name) =>
        Attributes.TryGetValue(name, out JsonElement value) && value.ValueKind == JsonValueKind.String
            ? value.GetString()
            : null;
}

internal static class AdfJsonValue {
    internal static JsonElement Create(object? value) {
        if (value is JsonElement element) return element.Clone();
        if (value is JsonDocument document) return document.RootElement.Clone();

        JsonNode? node = CreateNode(value);
        using var buffer = new MemoryStream();
        using (var writer = new Utf8JsonWriter(buffer)) {
            if (node == null) writer.WriteNullValue();
            else node.WriteTo(writer);
        }
        using JsonDocument parsed = JsonDocument.Parse(buffer.ToArray());
        return parsed.RootElement.Clone();
    }

    private static JsonNode? CreateNode(object? value) {
        if (value == null) return null;
        if (value is JsonNode node) return node.DeepClone();
        if (value is string text) return JsonValue.Create(text);
        if (value is bool boolean) return JsonValue.Create(boolean);
        if (value is byte byteValue) return JsonValue.Create(byteValue);
        if (value is sbyte signedByteValue) return JsonValue.Create(signedByteValue);
        if (value is short shortValue) return JsonValue.Create(shortValue);
        if (value is ushort unsignedShortValue) return JsonValue.Create(unsignedShortValue);
        if (value is int intValue) return JsonValue.Create(intValue);
        if (value is uint unsignedIntValue) return JsonValue.Create(unsignedIntValue);
        if (value is long longValue) return JsonValue.Create(longValue);
        if (value is ulong unsignedLongValue) return JsonValue.Create(unsignedLongValue);
        if (value is float floatValue) return JsonValue.Create(floatValue);
        if (value is double doubleValue) return JsonValue.Create(doubleValue);
        if (value is decimal decimalValue) return JsonValue.Create(decimalValue);
        if (value is char character) return JsonValue.Create(character.ToString());
        if (value is Guid guid) return JsonValue.Create(guid);
        if (value is DateTime dateTime) return JsonValue.Create(dateTime);
        if (value is DateTimeOffset dateTimeOffset) return JsonValue.Create(dateTimeOffset);
        if (value is Uri uri) return JsonValue.Create(uri.ToString());

        if (value is IReadOnlyDictionary<string, object?> readOnlyDictionary) {
            var result = new JsonObject();
            foreach (KeyValuePair<string, object?> entry in readOnlyDictionary) {
                result[entry.Key] = CreateNode(entry.Value);
            }
            return result;
        }

        if (value is IDictionary dictionary) {
            var result = new JsonObject();
            foreach (DictionaryEntry entry in dictionary) {
                if (entry.Key is not string key) throw new NotSupportedException("ADF JSON object keys must be strings.");
                result[key] = CreateNode(entry.Value);
            }
            return result;
        }

        if (value is IEnumerable sequence) {
            var result = new JsonArray();
            foreach (object? item in sequence) result.Add(CreateNode(item));
            return result;
        }

        return CreateNodeWithRuntimeSerializer(value);
    }

    [UnconditionalSuppressMessage("Trimming", "IL2026",
        Justification = "This compatibility branch is reached only on dynamic-code runtimes. NativeAOT callers use JSON-compatible scalar/collection values or the JsonTypeInfo overload.")]
    [UnconditionalSuppressMessage("AOT", "IL3050",
        Justification = "This compatibility branch is reached only on dynamic-code runtimes. NativeAOT callers use JSON-compatible scalar/collection values or the JsonTypeInfo overload.")]
    private static JsonNode? CreateNodeWithRuntimeSerializer(object value) {
#if NET5_0_OR_GREATER
        if (!RuntimeFeature.IsDynamicCodeSupported) {
            throw new NotSupportedException($"ADF attribute type '{value.GetType().FullName}' needs the SetAttribute overload with source-generated JsonTypeInfo metadata in NativeAOT applications.");
        }
#endif
        return JsonSerializer.SerializeToNode(value, value.GetType());
    }
}
