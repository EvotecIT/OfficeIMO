using System.IO;
using System.Text;
using System.Text.Json;
using System.Threading;

namespace OfficeIMO.Adf;

internal static class AdfJsonSerializer {
    private enum PropertySet { Root, Node, Mark }

    internal static AdfDocument Parse(string json, AdfProcessingOptions? options = null) {
        if (string.IsNullOrWhiteSpace(json)) throw new ArgumentException("ADF JSON is required.", nameof(json));
        options ??= new AdfProcessingOptions();
        AdfGraphGuard.Input(json, options);

        using JsonDocument source = JsonDocument.Parse(json, new JsonDocumentOptions { MaxDepth = options.MaxDepth * 3 + 8 });
        if (source.RootElement.ValueKind != JsonValueKind.Object) {
            throw new FormatException("An ADF document must be a JSON object.");
        }

        JsonElement root = source.RootElement;
        var path = new StringBuilder("$");
        var document = new AdfDocument {
            Version = ReadRequiredInt32(root, "version", "$"),
            Type = ReadRequiredString(root, "type", path),
        };

        if (!root.TryGetProperty("content", out JsonElement content) || content.ValueKind != JsonValueKind.Array) {
            throw new FormatException("ADF root property 'content' must be an array.");
        }

        int index = 0;
        foreach (JsonElement element in content.EnumerateArray()) {
            int pathLength = path.Length;
            path.Append(".content[").Append(index).Append(']');
            document.Content.Add(ReadNode(element, path, options));
            path.Length = pathLength;
            index++;
        }

        CopyExtensionProperties(root, document);
        AdfGraphSafety.EnsureSafe(document, options);
        return document;
    }

    internal static string Serialize(AdfDocument document, bool indented, AdfProcessingOptions? options = null) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        options ??= new AdfProcessingOptions();
        AdfGraphSafety.EnsureSafe(document, options);
        CancellationToken cancellationToken = options.CancellationToken;

        using var stream = new AdfJsonOutputStream(options);
        using (var writer = new Utf8JsonWriter(stream, new JsonWriterOptions { Indented = indented, MaxDepth = options.MaxDepth * 3 + 8 })) {
            writer.WriteStartObject();
            writer.WriteNumber("version", document.Version);
            writer.WriteString("type", document.Type);
            writer.WritePropertyName("content");
            writer.WriteStartArray();
            foreach (AdfNode node in document.ContentItems) WriteNode(writer, node, cancellationToken);
            writer.WriteEndArray();
            WriteExtensionProperties(writer, document.ExtensionItems, PropertySet.Root, cancellationToken);
            writer.WriteEndObject();
        }

        return Encoding.UTF8.GetString(stream.GetBuffer(), 0, checked((int)stream.Length));
    }

    private static AdfNode ReadNode(JsonElement element, StringBuilder path, AdfProcessingOptions options) {
        options.CancellationToken.ThrowIfCancellationRequested();
        if (element.ValueKind != JsonValueKind.Object) throw FormatException(path, " must be an object.");
        var node = new AdfNode(ReadRequiredString(element, "type", path));
        JsonElement text = default;
        JsonElement attributes = default;
        JsonElement content = default;
        JsonElement marks = default;
        bool hasText = false;
        bool hasAttributes = false;
        bool hasContent = false;
        bool hasMarks = false;
        foreach (JsonProperty property in element.EnumerateObject()) {
            switch (property.Name) {
                case "type":
                    break;
                case "text":
                    text = property.Value;
                    hasText = true;
                    break;
                case "attrs":
                    attributes = property.Value;
                    hasAttributes = true;
                    break;
                case "content":
                    content = property.Value;
                    hasContent = true;
                    break;
                case "marks":
                    marks = property.Value;
                    hasMarks = true;
                    break;
                default:
                    node.AddExtension(property.Name, property.Value.Clone());
                    break;
            }
        }
        if (hasText) {
            if (text.ValueKind != JsonValueKind.String) throw FormatException(path, ".text must be a string.");
            node.Text = text.GetString();
        }
        if (hasAttributes) {
            if (attributes.ValueKind != JsonValueKind.Object) throw FormatException(path, ".attrs must be an object.");
            node.IncludeEmptyAttributes = true;
            foreach (JsonProperty attribute in attributes.EnumerateObject()) {
                node.AddAttribute(attribute.Name, attribute.Value.Clone());
            }
        }
        if (hasContent) {
            if (content.ValueKind != JsonValueKind.Array) throw FormatException(path, ".content must be an array.");
            node.IncludeEmptyContent = true;
            int childIndex = 0;
            foreach (JsonElement child in content.EnumerateArray()) {
                int pathLength = path.Length;
                path.Append(".content[").Append(childIndex).Append(']');
                node.Content.Add(ReadNode(child, path, options));
                path.Length = pathLength;
                childIndex++;
            }
        }
        if (hasMarks) {
            if (marks.ValueKind != JsonValueKind.Array) throw FormatException(path, ".marks must be an array.");
            node.IncludeEmptyMarks = true;
            int markIndex = 0;
            foreach (JsonElement mark in marks.EnumerateArray()) {
                options.CancellationToken.ThrowIfCancellationRequested();
                int pathLength = path.Length;
                path.Append(".marks[").Append(markIndex).Append(']');
                node.Marks.Add(ReadMark(mark, path));
                path.Length = pathLength;
                markIndex++;
            }
        }
        return node;
    }

    private static AdfMark ReadMark(JsonElement element, StringBuilder path) {
        if (element.ValueKind != JsonValueKind.Object) throw FormatException(path, " must be an object.");
        var mark = new AdfMark(ReadRequiredString(element, "type", path));
        JsonElement attributes = default;
        bool hasAttributes = false;
        foreach (JsonProperty property in element.EnumerateObject()) {
            switch (property.Name) {
                case "type":
                    break;
                case "attrs":
                    attributes = property.Value;
                    hasAttributes = true;
                    break;
                default:
                    mark.AddExtension(property.Name, property.Value.Clone());
                    break;
            }
        }
        if (hasAttributes) {
            if (attributes.ValueKind != JsonValueKind.Object) throw FormatException(path, ".attrs must be an object.");
            mark.IncludeEmptyAttributes = true;
            foreach (JsonProperty attribute in attributes.EnumerateObject()) {
                mark.AddAttribute(attribute.Name, attribute.Value.Clone());
            }
        }
        return mark;
    }

    private static void WriteNode(Utf8JsonWriter writer, AdfNode node, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (node == null) throw new InvalidOperationException("ADF content cannot contain null nodes.");
        writer.WriteStartObject();
        writer.WriteString("type", node.Type);
        if (node.AttributeItems.Count > 0 || node.IncludeEmptyAttributes) WriteObject(writer, "attrs", node.AttributeItems, cancellationToken);
        if (node.Text != null) writer.WriteString("text", node.Text);
        if (node.MarkItems.Count > 0 || node.IncludeEmptyMarks) {
            writer.WritePropertyName("marks");
            writer.WriteStartArray();
            foreach (AdfMark mark in node.MarkItems) WriteMark(writer, mark, cancellationToken);
            writer.WriteEndArray();
        }
        if (node.ContentItems.Count > 0 || node.IncludeEmptyContent || AdfNodeShape.RequiresContent(node.Type)) {
            writer.WritePropertyName("content");
            writer.WriteStartArray();
            foreach (AdfNode child in node.ContentItems) WriteNode(writer, child, cancellationToken);
            writer.WriteEndArray();
        }
        WriteExtensionProperties(writer, node.ExtensionItems, PropertySet.Node, cancellationToken);
        writer.WriteEndObject();
    }

    private static void WriteMark(Utf8JsonWriter writer, AdfMark mark, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        writer.WriteStartObject();
        writer.WriteString("type", mark.Type);
        if (mark.AttributeItems.Count > 0 || mark.IncludeEmptyAttributes) WriteObject(writer, "attrs", mark.AttributeItems, cancellationToken);
        WriteExtensionProperties(writer, mark.ExtensionItems, PropertySet.Mark, cancellationToken);
        writer.WriteEndObject();
    }

    private static void WriteObject(Utf8JsonWriter writer, string name, IReadOnlyDictionary<string, JsonElement> values, CancellationToken cancellationToken) {
        writer.WritePropertyName(name);
        writer.WriteStartObject();
        foreach (KeyValuePair<string, JsonElement> value in values) {
            cancellationToken.ThrowIfCancellationRequested();
            writer.WritePropertyName(value.Key);
            value.Value.WriteTo(writer);
        }
        writer.WriteEndObject();
    }

    private static void CopyExtensionProperties(JsonElement source, AdfDocument target) {
        foreach (JsonProperty property in source.EnumerateObject()) {
            if (!IsKnownProperty(property.Name, PropertySet.Root)) target.AddExtension(property.Name, property.Value.Clone());
        }
    }

    private static void WriteExtensionProperties(Utf8JsonWriter writer, IReadOnlyDictionary<string, JsonElement> values, PropertySet propertySet, CancellationToken cancellationToken) {
        foreach (KeyValuePair<string, JsonElement> value in values) {
            cancellationToken.ThrowIfCancellationRequested();
            if (IsKnownProperty(value.Key, propertySet)) continue;
            writer.WritePropertyName(value.Key);
            value.Value.WriteTo(writer);
        }
    }

    private static bool IsKnownProperty(string name, PropertySet propertySet) => propertySet switch {
        PropertySet.Root => name is "version" or "type" or "content",
        PropertySet.Node => name is "type" or "text" or "attrs" or "content" or "marks",
        PropertySet.Mark => name is "type" or "attrs",
        _ => false,
    };

    private static string ReadRequiredString(JsonElement element, string name, StringBuilder path) {
        if (!element.TryGetProperty(name, out JsonElement value) || value.ValueKind != JsonValueKind.String || string.IsNullOrWhiteSpace(value.GetString())) {
            throw FormatException(path, "." + name + " must be a non-empty string.");
        }
        return value.GetString()!;
    }

    private static FormatException FormatException(StringBuilder path, string suffix) =>
        new(path.ToString() + suffix);

    private static int ReadRequiredInt32(JsonElement element, string name, string path) {
        if (!element.TryGetProperty(name, out JsonElement value) || !AdfJsonValue.TryGetInt32(value, out int result)) {
            throw new FormatException(path + "." + name + " must be an integer.");
        }
        return result;
    }
}
