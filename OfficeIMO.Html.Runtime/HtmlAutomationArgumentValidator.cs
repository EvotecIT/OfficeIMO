using System.Text.Json;

namespace OfficeIMO.Html.Runtime;

// Validates the fixed built-in tool schema vocabulary. This is not a general JSON Schema engine.
internal static class HtmlAutomationArgumentValidator {
    internal static JsonElement Normalize(JsonElement arguments, JsonElement schema) {
        using var stream = new MemoryStream();
        using (var writer = new Utf8JsonWriter(stream)) Write(writer, arguments, schema, "$arguments");
        using JsonDocument document = JsonDocument.Parse(stream.ToArray());
        return document.RootElement.Clone();
    }

    private static void Write(Utf8JsonWriter writer, JsonElement value, JsonElement schema, string path) {
        string type = schema.GetProperty("type").GetString()!;
        bool matches = type switch {
            "object" => value.ValueKind == JsonValueKind.Object,
            "array" => value.ValueKind == JsonValueKind.Array,
            "string" => value.ValueKind == JsonValueKind.String,
            "boolean" => value.ValueKind is JsonValueKind.True or JsonValueKind.False,
            "integer" => value.ValueKind == JsonValueKind.Number && value.TryGetInt64(out _),
            _ => false
        };
        if (!matches) throw Invalid(path, "has an invalid type");
        if (type == "object") {
            JsonElement properties = schema.GetProperty("properties");
            var required = schema.TryGetProperty("required", out JsonElement requirements)
                ? requirements.EnumerateArray().Select(item => item.GetString()!).ToHashSet(StringComparer.Ordinal)
                : new HashSet<string>(StringComparer.Ordinal);
            var seen = new HashSet<string>(StringComparer.Ordinal);
            writer.WriteStartObject();
            foreach (JsonProperty property in value.EnumerateObject()) {
                if (!seen.Add(property.Name)) throw Invalid(path + "." + property.Name, "is duplicate");
                if (!properties.TryGetProperty(property.Name, out JsonElement child))
                    throw Invalid(path + "." + property.Name, "is unknown");
                if (property.Value.ValueKind == JsonValueKind.Null && !required.Contains(property.Name)) continue;
                writer.WritePropertyName(property.Name);
                Write(writer, property.Value, child, path + "." + property.Name);
            }
            if (required.Any(name => !seen.Contains(name))) throw Invalid(path, "is missing a required property");
            writer.WriteEndObject();
        } else if (type == "array") {
            writer.WriteStartArray();
            int index = 0;
            foreach (JsonElement item in value.EnumerateArray()) Write(writer, item, schema.GetProperty("items"), path + "[" + index++ + "]");
            writer.WriteEndArray();
        } else {
            if (schema.TryGetProperty("enum", out JsonElement allowed)
                && !allowed.EnumerateArray().Any(item => item.GetString() == value.GetString()))
                throw Invalid(path, "is outside the allowed values");
            if (type == "integer") {
                long number = value.GetInt64();
                if (schema.TryGetProperty("minimum", out JsonElement minimum) && number < minimum.GetInt64()
                    || schema.TryGetProperty("maximum", out JsonElement maximum) && number > maximum.GetInt64())
                    throw Invalid(path, "is outside the allowed range");
            }
            if (schema.TryGetProperty("format", out JsonElement format) && format.GetString() == "uri"
                && !Uri.TryCreate(value.GetString(), UriKind.Absolute, out _)) throw Invalid(path, "requires an absolute URI");
            value.WriteTo(writer);
        }
    }

    private static ArgumentException Invalid(string path, string reason) => new($"Tool argument '{path}' {reason}.");
}
