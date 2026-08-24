using System.IO;
using System.Text;
using System.Text.Json;

namespace OfficeIMO.Adf;

internal static class AdfJsonSerializer {
    private enum PropertySet { Root, Node, Mark }

    internal static AdfDocument Parse(string json) {
        if (string.IsNullOrWhiteSpace(json)) throw new ArgumentException("ADF JSON is required.", nameof(json));

        using JsonDocument source = JsonDocument.Parse(json);
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
            document.Content.Add(ReadNode(element, path));
            path.Length = pathLength;
            index++;
        }

        CopyExtensionProperties(root, document);
        return document;
    }

    internal static string Serialize(AdfDocument document, bool indented) {
        if (document == null) throw new ArgumentNullException(nameof(document));

        using var stream = new MemoryStream();
        using (var writer = new Utf8JsonWriter(stream, new JsonWriterOptions { Indented = indented })) {
            writer.WriteStartObject();
            writer.WriteNumber("version", document.Version);
            writer.WriteString("type", document.Type);
            writer.WritePropertyName("content");
            writer.WriteStartArray();
            foreach (AdfNode node in document.ContentItems) WriteNode(writer, node);
            writer.WriteEndArray();
            WriteExtensionProperties(writer, document.ExtensionItems, PropertySet.Root);
            writer.WriteEndObject();
        }

        return Encoding.UTF8.GetString(stream.GetBuffer(), 0, checked((int)stream.Length));
    }

    private static AdfNode ReadNode(JsonElement element, StringBuilder path) {
        if (element.ValueKind != JsonValueKind.Object) throw FormatException(path, " must be an object.");
        var node = new AdfNode(ReadRequiredString(element, "type", path));

        if (element.TryGetProperty("text", out JsonElement text)) {
            if (text.ValueKind != JsonValueKind.String) throw FormatException(path, ".text must be a string.");
            node.Text = text.GetString();
        }

        if (element.TryGetProperty("attrs", out JsonElement attributes)) {
            if (attributes.ValueKind != JsonValueKind.Object) throw FormatException(path, ".attrs must be an object.");
            foreach (JsonProperty property in attributes.EnumerateObject()) {
                node.AddAttribute(property.Name, property.Value.Clone());
            }
        }

        if (element.TryGetProperty("content", out JsonElement content)) {
            if (content.ValueKind != JsonValueKind.Array) throw FormatException(path, ".content must be an array.");
            int index = 0;
            foreach (JsonElement child in content.EnumerateArray()) {
                int pathLength = path.Length;
                path.Append(".content[").Append(index).Append(']');
                node.Content.Add(ReadNode(child, path));
                path.Length = pathLength;
                index++;
            }
        }

        if (element.TryGetProperty("marks", out JsonElement marks)) {
            if (marks.ValueKind != JsonValueKind.Array) throw FormatException(path, ".marks must be an array.");
            int index = 0;
            foreach (JsonElement mark in marks.EnumerateArray()) {
                int pathLength = path.Length;
                path.Append(".marks[").Append(index).Append(']');
                node.Marks.Add(ReadMark(mark, path));
                path.Length = pathLength;
                index++;
            }
        }

        CopyExtensionProperties(element, node);
        return node;
    }

    private static AdfMark ReadMark(JsonElement element, StringBuilder path) {
        if (element.ValueKind != JsonValueKind.Object) throw FormatException(path, " must be an object.");
        var mark = new AdfMark(ReadRequiredString(element, "type", path));
        if (element.TryGetProperty("attrs", out JsonElement attributes)) {
            if (attributes.ValueKind != JsonValueKind.Object) throw FormatException(path, ".attrs must be an object.");
            foreach (JsonProperty property in attributes.EnumerateObject()) {
                mark.AddAttribute(property.Name, property.Value.Clone());
            }
        }
        CopyExtensionProperties(element, mark);
        return mark;
    }

    private static void WriteNode(Utf8JsonWriter writer, AdfNode node) {
        if (node == null) throw new InvalidOperationException("ADF content cannot contain null nodes.");
        writer.WriteStartObject();
        writer.WriteString("type", node.Type);
        if (node.AttributeItems.Count > 0) WriteObject(writer, "attrs", node.AttributeItems);
        if (node.Text != null) writer.WriteString("text", node.Text);
        if (node.MarkItems.Count > 0) {
            writer.WritePropertyName("marks");
            writer.WriteStartArray();
            foreach (AdfMark mark in node.MarkItems) WriteMark(writer, mark);
            writer.WriteEndArray();
        }
        if (node.ContentItems.Count > 0) {
            writer.WritePropertyName("content");
            writer.WriteStartArray();
            foreach (AdfNode child in node.ContentItems) WriteNode(writer, child);
            writer.WriteEndArray();
        }
        WriteExtensionProperties(writer, node.ExtensionItems, PropertySet.Node);
        writer.WriteEndObject();
    }

    private static void WriteMark(Utf8JsonWriter writer, AdfMark mark) {
        writer.WriteStartObject();
        writer.WriteString("type", mark.Type);
        if (mark.AttributeItems.Count > 0) WriteObject(writer, "attrs", mark.AttributeItems);
        WriteExtensionProperties(writer, mark.ExtensionItems, PropertySet.Mark);
        writer.WriteEndObject();
    }

    private static void WriteObject(Utf8JsonWriter writer, string name, IReadOnlyDictionary<string, JsonElement> values) {
        writer.WritePropertyName(name);
        writer.WriteStartObject();
        foreach (KeyValuePair<string, JsonElement> value in values) {
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

    private static void CopyExtensionProperties(JsonElement source, AdfNode target) {
        foreach (JsonProperty property in source.EnumerateObject()) {
            if (!IsKnownProperty(property.Name, PropertySet.Node)) target.AddExtension(property.Name, property.Value.Clone());
        }
    }

    private static void CopyExtensionProperties(JsonElement source, AdfMark target) {
        foreach (JsonProperty property in source.EnumerateObject()) {
            if (!IsKnownProperty(property.Name, PropertySet.Mark)) target.AddExtension(property.Name, property.Value.Clone());
        }
    }

    private static void WriteExtensionProperties(Utf8JsonWriter writer, IReadOnlyDictionary<string, JsonElement> values, PropertySet propertySet) {
        foreach (KeyValuePair<string, JsonElement> value in values) {
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
        if (!element.TryGetProperty(name, out JsonElement value) || value.ValueKind != JsonValueKind.Number || !value.TryGetInt32(out int result)) {
            throw new FormatException(path + "." + name + " must be an integer.");
        }
        return result;
    }
}
