using System.IO;
using System.Text;
using System.Text.Json;

namespace OfficeIMO.Adf;

internal static class AdfJsonSerializer {
    internal static AdfDocument Parse(string json) {
        if (string.IsNullOrWhiteSpace(json)) throw new ArgumentException("ADF JSON is required.", nameof(json));

        using JsonDocument source = JsonDocument.Parse(json);
        if (source.RootElement.ValueKind != JsonValueKind.Object) {
            throw new FormatException("An ADF document must be a JSON object.");
        }

        JsonElement root = source.RootElement;
        var document = new AdfDocument {
            Version = ReadRequiredInt32(root, "version", "$"),
            Type = ReadRequiredString(root, "type", "$"),
        };

        if (!root.TryGetProperty("content", out JsonElement content) || content.ValueKind != JsonValueKind.Array) {
            throw new FormatException("ADF root property 'content' must be an array.");
        }

        int index = 0;
        foreach (JsonElement element in content.EnumerateArray()) {
            document.Content.Add(ReadNode(element, "$.content[" + index + "]"));
            index++;
        }

        CopyExtensionProperties(root, document.ExtensionData, "version", "type", "content");
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
            foreach (AdfNode node in document.Content) WriteNode(writer, node);
            writer.WriteEndArray();
            WriteExtensionProperties(writer, document.ExtensionData, "version", "type", "content");
            writer.WriteEndObject();
        }

        return Encoding.UTF8.GetString(stream.ToArray());
    }

    private static AdfNode ReadNode(JsonElement element, string path) {
        if (element.ValueKind != JsonValueKind.Object) throw new FormatException(path + " must be an object.");
        var node = new AdfNode(ReadRequiredString(element, "type", path));

        if (element.TryGetProperty("text", out JsonElement text)) {
            if (text.ValueKind != JsonValueKind.String) throw new FormatException(path + ".text must be a string.");
            node.Text = text.GetString();
        }

        if (element.TryGetProperty("attrs", out JsonElement attributes)) {
            ReadObjectProperties(attributes, node.Attributes, path + ".attrs");
        }

        if (element.TryGetProperty("content", out JsonElement content)) {
            if (content.ValueKind != JsonValueKind.Array) throw new FormatException(path + ".content must be an array.");
            int index = 0;
            foreach (JsonElement child in content.EnumerateArray()) {
                node.Content.Add(ReadNode(child, path + ".content[" + index + "]"));
                index++;
            }
        }

        if (element.TryGetProperty("marks", out JsonElement marks)) {
            if (marks.ValueKind != JsonValueKind.Array) throw new FormatException(path + ".marks must be an array.");
            int index = 0;
            foreach (JsonElement mark in marks.EnumerateArray()) {
                node.Marks.Add(ReadMark(mark, path + ".marks[" + index + "]"));
                index++;
            }
        }

        CopyExtensionProperties(element, node.ExtensionData, "type", "text", "attrs", "content", "marks");
        return node;
    }

    private static AdfMark ReadMark(JsonElement element, string path) {
        if (element.ValueKind != JsonValueKind.Object) throw new FormatException(path + " must be an object.");
        var mark = new AdfMark(ReadRequiredString(element, "type", path));
        if (element.TryGetProperty("attrs", out JsonElement attributes)) {
            ReadObjectProperties(attributes, mark.Attributes, path + ".attrs");
        }
        CopyExtensionProperties(element, mark.ExtensionData, "type", "attrs");
        return mark;
    }

    private static void WriteNode(Utf8JsonWriter writer, AdfNode node) {
        if (node == null) throw new InvalidOperationException("ADF content cannot contain null nodes.");
        writer.WriteStartObject();
        writer.WriteString("type", node.Type);
        if (node.Attributes.Count > 0) WriteObject(writer, "attrs", node.Attributes);
        if (node.Text != null) writer.WriteString("text", node.Text);
        if (node.Marks.Count > 0) {
            writer.WritePropertyName("marks");
            writer.WriteStartArray();
            foreach (AdfMark mark in node.Marks) WriteMark(writer, mark);
            writer.WriteEndArray();
        }
        if (node.Content.Count > 0) {
            writer.WritePropertyName("content");
            writer.WriteStartArray();
            foreach (AdfNode child in node.Content) WriteNode(writer, child);
            writer.WriteEndArray();
        }
        WriteExtensionProperties(writer, node.ExtensionData, "type", "text", "attrs", "content", "marks");
        writer.WriteEndObject();
    }

    private static void WriteMark(Utf8JsonWriter writer, AdfMark mark) {
        writer.WriteStartObject();
        writer.WriteString("type", mark.Type);
        if (mark.Attributes.Count > 0) WriteObject(writer, "attrs", mark.Attributes);
        WriteExtensionProperties(writer, mark.ExtensionData, "type", "attrs");
        writer.WriteEndObject();
    }

    private static void ReadObjectProperties(JsonElement element, IDictionary<string, JsonElement> target, string path) {
        if (element.ValueKind != JsonValueKind.Object) throw new FormatException(path + " must be an object.");
        foreach (JsonProperty property in element.EnumerateObject()) target[property.Name] = property.Value.Clone();
    }

    private static void WriteObject(Utf8JsonWriter writer, string name, IDictionary<string, JsonElement> values) {
        writer.WritePropertyName(name);
        writer.WriteStartObject();
        foreach (KeyValuePair<string, JsonElement> value in values) {
            writer.WritePropertyName(value.Key);
            value.Value.WriteTo(writer);
        }
        writer.WriteEndObject();
    }

    private static void CopyExtensionProperties(JsonElement source, IDictionary<string, JsonElement> target, params string[] knownNames) {
        var known = new HashSet<string>(knownNames, StringComparer.Ordinal);
        foreach (JsonProperty property in source.EnumerateObject()) {
            if (!known.Contains(property.Name)) target[property.Name] = property.Value.Clone();
        }
    }

    private static void WriteExtensionProperties(Utf8JsonWriter writer, IDictionary<string, JsonElement> values, params string[] knownNames) {
        var known = new HashSet<string>(knownNames, StringComparer.Ordinal);
        foreach (KeyValuePair<string, JsonElement> value in values) {
            if (known.Contains(value.Key)) continue;
            writer.WritePropertyName(value.Key);
            value.Value.WriteTo(writer);
        }
    }

    private static string ReadRequiredString(JsonElement element, string name, string path) {
        if (!element.TryGetProperty(name, out JsonElement value) || value.ValueKind != JsonValueKind.String || string.IsNullOrWhiteSpace(value.GetString())) {
            throw new FormatException(path + "." + name + " must be a non-empty string.");
        }
        return value.GetString()!;
    }

    private static int ReadRequiredInt32(JsonElement element, string name, string path) {
        if (!element.TryGetProperty(name, out JsonElement value) || value.ValueKind != JsonValueKind.Number || !value.TryGetInt32(out int result)) {
            throw new FormatException(path + "." + name + " must be an integer.");
        }
        return result;
    }
}
