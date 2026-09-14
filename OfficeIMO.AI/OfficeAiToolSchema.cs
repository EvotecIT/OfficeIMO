using System.Text.Json;

namespace OfficeIMO.AI;

/// <summary>Owned, dependency-free validation for the bounded JSON Schema subset accepted by tool planning.</summary>
internal static class OfficeAiToolSchema {
    private static readonly HashSet<string> AllowedKeywords = new(StringComparer.Ordinal) {
        "$schema", "type", "description", "properties", "required", "additionalProperties", "items",
        "enum", "const", "anyOf", "minLength", "maxLength", "minimum", "maximum", "minItems", "maxItems", "format"
    };

    internal static void ValidateDefinition(JsonElement schema) {
        ValidateSchema(schema, isRoot: true, depth: 1);
        if (!AllowsType(schema, "object"))
            throw new ArgumentException("A tool input schema must describe a JSON object.", nameof(schema));
    }

    internal static JsonElement NormalizeAndValidate(JsonElement value, JsonElement schema) {
        using var stream = new MemoryStream();
        using (var writer = new Utf8JsonWriter(stream)) WriteValue(writer, value, schema, "$", 1);
        using JsonDocument normalized = JsonDocument.Parse(stream.ToArray());
        return normalized.RootElement.Clone();
    }

    internal static void WriteStrict(Utf8JsonWriter writer, JsonElement schema, bool nullable) {
        if (nullable && !AllowsType(schema, "null")) {
            writer.WriteStartObject();
            writer.WritePropertyName("anyOf");
            writer.WriteStartArray();
            WriteStrict(writer, schema, nullable: false);
            writer.WriteStartObject(); writer.WriteString("type", "null"); writer.WriteEndObject();
            writer.WriteEndArray();
            writer.WriteEndObject();
            return;
        }
        if (schema.TryGetProperty("anyOf", out JsonElement choices)) {
            writer.WriteStartObject(); writer.WritePropertyName("anyOf"); writer.WriteStartArray();
            foreach (JsonElement choice in choices.EnumerateArray()) WriteStrict(writer, choice, nullable: false);
            writer.WriteEndArray(); writer.WriteEndObject();
            return;
        }

        writer.WriteStartObject();
        foreach (string keyword in new[] { "type", "description", "enum", "const" })
            if (schema.TryGetProperty(keyword, out JsonElement item)) { writer.WritePropertyName(keyword); item.WriteTo(writer); }
        string? type = PrimaryType(schema);
        if (type == "object") {
            JsonElement properties = schema.GetProperty("properties");
            writer.WritePropertyName("properties"); writer.WriteStartObject();
            var names = new List<string>();
            var required = RequiredNames(schema);
            foreach (JsonProperty property in properties.EnumerateObject()) {
                names.Add(property.Name);
                writer.WritePropertyName(property.Name);
                WriteStrict(writer, property.Value, nullable: !required.Contains(property.Name));
            }
            writer.WriteEndObject();
            writer.WritePropertyName("required"); writer.WriteStartArray();
            foreach (string name in names) writer.WriteStringValue(name);
            writer.WriteEndArray();
            writer.WriteBoolean("additionalProperties", false);
        } else if (type == "array") {
            writer.WritePropertyName("items");
            WriteStrict(writer, schema.GetProperty("items"), nullable: false);
        }
        writer.WriteEndObject();
    }

    private static void ValidateSchema(JsonElement schema, bool isRoot, int depth) {
        if (depth > 16 || schema.ValueKind != JsonValueKind.Object)
            throw new ArgumentException("Tool schemas must contain objects with at most 16 nested levels.", nameof(schema));
        var seen = new HashSet<string>(StringComparer.Ordinal);
        foreach (JsonProperty property in schema.EnumerateObject()) {
            if (!seen.Add(property.Name) || !AllowedKeywords.Contains(property.Name))
                throw new ArgumentException($"Unsupported or duplicate tool-schema keyword '{property.Name}'.", nameof(schema));
            if (property.Name == "$schema" && !isRoot)
                throw new ArgumentException("$schema is allowed only at the tool-schema root.", nameof(schema));
        }
        bool hasAnyOf = schema.TryGetProperty("anyOf", out JsonElement anyOf);
        if (hasAnyOf) {
            if (schema.TryGetProperty("type", out _) || anyOf.ValueKind != JsonValueKind.Array || anyOf.GetArrayLength() < 1)
                throw new ArgumentException("anyOf must contain one or more schemas and cannot be combined with type.", nameof(schema));
            foreach (JsonElement choice in anyOf.EnumerateArray()) ValidateSchema(choice, isRoot: false, depth + 1);
            return;
        }
        HashSet<string> types = ReadTypes(schema);
        if (types.Count == 0) throw new ArgumentException("Every tool schema must declare type or anyOf.", nameof(schema));
        if (types.Contains("object")) ValidateObjectSchema(schema, depth);
        if (types.Contains("array")) {
            if (!schema.TryGetProperty("items", out JsonElement items))
                throw new ArgumentException("Array tool schemas require items.", nameof(schema));
            ValidateSchema(items, isRoot: false, depth + 1);
        }
        ValidateNonNegativeInteger(schema, "minLength"); ValidateNonNegativeInteger(schema, "maxLength");
        ValidateNonNegativeInteger(schema, "minItems"); ValidateNonNegativeInteger(schema, "maxItems");
        ValidateNumber(schema, "minimum"); ValidateNumber(schema, "maximum");
        if (schema.TryGetProperty("format", out JsonElement format)
            && (format.ValueKind != JsonValueKind.String || format.GetString() != "uri"))
            throw new ArgumentException("The owned tool-schema subset supports only the uri format.", nameof(schema));
        if (schema.TryGetProperty("enum", out JsonElement values) && (values.ValueKind != JsonValueKind.Array || values.GetArrayLength() == 0))
            throw new ArgumentException("A tool-schema enum must contain at least one value.", nameof(schema));
    }

    private static void ValidateObjectSchema(JsonElement schema, int depth) {
        if (!schema.TryGetProperty("properties", out JsonElement properties) || properties.ValueKind != JsonValueKind.Object)
            throw new ArgumentException("Object tool schemas require a properties object.", nameof(schema));
        if (!schema.TryGetProperty("additionalProperties", out JsonElement additional)
            || additional.ValueKind != JsonValueKind.False)
            throw new ArgumentException("Object tool schemas must set additionalProperties to false.", nameof(schema));
        var names = new HashSet<string>(StringComparer.Ordinal);
        foreach (JsonProperty property in properties.EnumerateObject()) {
            if (!names.Add(property.Name)) throw new ArgumentException($"Duplicate tool-schema property '{property.Name}'.", nameof(schema));
            ValidateSchema(property.Value, isRoot: false, depth + 1);
        }
        HashSet<string> required = RequiredNames(schema);
        if (required.Any(name => !names.Contains(name)))
            throw new ArgumentException("Every required tool-schema property must be declared.", nameof(schema));
    }

    private static HashSet<string> ReadTypes(JsonElement schema) {
        var result = new HashSet<string>(StringComparer.Ordinal);
        if (!schema.TryGetProperty("type", out JsonElement type)) return result;
        if (type.ValueKind == JsonValueKind.String) AddType(type.GetString(), result);
        else if (type.ValueKind == JsonValueKind.Array) {
            foreach (JsonElement item in type.EnumerateArray()) {
                if (item.ValueKind != JsonValueKind.String) throw new ArgumentException("Tool-schema types must be strings.", nameof(schema));
                AddType(item.GetString(), result);
            }
        } else throw new ArgumentException("Tool-schema type must be a string or string array.", nameof(schema));
        return result;
    }

    private static void AddType(string? value, HashSet<string> types) {
        if (value is not ("object" or "array" or "string" or "integer" or "number" or "boolean" or "null") || !types.Add(value))
            throw new ArgumentException($"Unsupported or duplicate tool-schema type '{value}'.", nameof(types));
    }

    private static HashSet<string> RequiredNames(JsonElement schema) {
        var result = new HashSet<string>(StringComparer.Ordinal);
        if (!schema.TryGetProperty("required", out JsonElement required)) return result;
        if (required.ValueKind != JsonValueKind.Array) throw new ArgumentException("Tool-schema required must be an array.", nameof(schema));
        foreach (JsonElement item in required.EnumerateArray()) {
            string? name = item.ValueKind == JsonValueKind.String ? item.GetString() : null;
            if (string.IsNullOrEmpty(name) || !result.Add(name))
                throw new ArgumentException("Tool-schema required names must be unique non-empty strings.", nameof(schema));
        }
        return result;
    }

    private static void WriteValue(Utf8JsonWriter writer, JsonElement value, JsonElement schema, string path, int depth) {
        if (depth > 64) throw Invalid(path, "exceeds the supported nesting depth");
        if (schema.TryGetProperty("anyOf", out JsonElement choices)) {
            foreach (JsonElement choice in choices.EnumerateArray()) {
                try {
                    using var candidate = new MemoryStream();
                    using (var candidateWriter = new Utf8JsonWriter(candidate)) WriteValue(candidateWriter, value, choice, path, depth);
                    using JsonDocument normalized = JsonDocument.Parse(candidate.ToArray());
                    normalized.RootElement.WriteTo(writer);
                    return;
                } catch (InvalidOperationException) { }
            }
            throw Invalid(path, "does not match any allowed schema");
        }
        if (!MatchesType(value, schema)) throw Invalid(path, "has the wrong JSON type");
        ValidateConstAndEnum(value, schema, path);
        switch (value.ValueKind) {
            case JsonValueKind.Object: WriteObject(writer, value, schema, path, depth); break;
            case JsonValueKind.Array: WriteArray(writer, value, schema, path, depth); break;
            case JsonValueKind.String: ValidateString(value.GetString()!, schema, path); value.WriteTo(writer); break;
            case JsonValueKind.Number: ValidateNumeric(value, schema, path); value.WriteTo(writer); break;
            default: value.WriteTo(writer); break;
        }
    }

    private static void WriteObject(Utf8JsonWriter writer, JsonElement value, JsonElement schema, string path, int depth) {
        JsonElement properties = schema.GetProperty("properties");
        HashSet<string> required = RequiredNames(schema);
        var seen = new HashSet<string>(StringComparer.Ordinal);
        writer.WriteStartObject();
        foreach (JsonProperty property in value.EnumerateObject()) {
            if (!seen.Add(property.Name)) throw Invalid(path, $"contains duplicate property '{property.Name}'");
            if (!properties.TryGetProperty(property.Name, out JsonElement propertySchema))
                throw Invalid(path, $"contains unknown property '{property.Name}'");
            if (property.Value.ValueKind == JsonValueKind.Null && !required.Contains(property.Name) && !AllowsType(propertySchema, "null"))
                continue;
            writer.WritePropertyName(property.Name);
            WriteValue(writer, property.Value, propertySchema, path + "." + property.Name, depth + 1);
        }
        foreach (string name in required)
            if (!seen.Contains(name)) throw Invalid(path, $"is missing required property '{name}'");
        writer.WriteEndObject();
    }

    private static void WriteArray(Utf8JsonWriter writer, JsonElement value, JsonElement schema, string path, int depth) {
        int count = value.GetArrayLength();
        CheckRange(count, schema, "minItems", "maxItems", path);
        writer.WriteStartArray();
        int index = 0;
        foreach (JsonElement item in value.EnumerateArray()) WriteValue(writer, item, schema.GetProperty("items"), $"{path}[{index++}]", depth + 1);
        writer.WriteEndArray();
    }

    private static void ValidateString(string value, JsonElement schema, string path) {
        CheckRange(value.Length, schema, "minLength", "maxLength", path);
        if (schema.TryGetProperty("format", out _) && !Uri.TryCreate(value, UriKind.Absolute, out _))
            throw Invalid(path, "must be an absolute URI");
    }

    private static void ValidateNumeric(JsonElement value, JsonElement schema, string path) {
        double number = value.GetDouble();
        if (!double.IsFinite(number)) throw Invalid(path, "must be finite");
        if (AllowsType(schema, "integer") && Math.Truncate(number) != number) throw Invalid(path, "must be an integer");
        if (schema.TryGetProperty("minimum", out JsonElement minimum) && number < minimum.GetDouble()) throw Invalid(path, "is below minimum");
        if (schema.TryGetProperty("maximum", out JsonElement maximum) && number > maximum.GetDouble()) throw Invalid(path, "is above maximum");
    }

    private static void ValidateConstAndEnum(JsonElement value, JsonElement schema, string path) {
        if (schema.TryGetProperty("const", out JsonElement constant) && !JsonEquals(value, constant)) throw Invalid(path, "does not match const");
        if (schema.TryGetProperty("enum", out JsonElement values)
            && !values.EnumerateArray().Any(item => JsonEquals(value, item))) throw Invalid(path, "is outside the allowed enum");
    }

    private static bool JsonEquals(JsonElement left, JsonElement right) =>
        left.ValueKind == right.ValueKind && left.GetRawText() == right.GetRawText();

    private static bool MatchesType(JsonElement value, JsonElement schema) => value.ValueKind switch {
        JsonValueKind.Object => AllowsType(schema, "object"), JsonValueKind.Array => AllowsType(schema, "array"),
        JsonValueKind.String => AllowsType(schema, "string"), JsonValueKind.Number => AllowsType(schema, "number") || AllowsType(schema, "integer"),
        JsonValueKind.True or JsonValueKind.False => AllowsType(schema, "boolean"), JsonValueKind.Null => AllowsType(schema, "null"), _ => false
    };

    private static bool AllowsType(JsonElement schema, string type) =>
        schema.TryGetProperty("anyOf", out JsonElement choices)
            ? choices.EnumerateArray().Any(choice => AllowsType(choice, type))
            : ReadTypes(schema).Contains(type);

    private static string? PrimaryType(JsonElement schema) {
        HashSet<string> types = ReadTypes(schema);
        return types.FirstOrDefault(type => type != "null");
    }

    private static void CheckRange(int value, JsonElement schema, string minimumName, string maximumName, string path) {
        if (schema.TryGetProperty(minimumName, out JsonElement minimum) && value < minimum.GetInt32()) throw Invalid(path, $"is below {minimumName}");
        if (schema.TryGetProperty(maximumName, out JsonElement maximum) && value > maximum.GetInt32()) throw Invalid(path, $"exceeds {maximumName}");
    }

    private static void ValidateNonNegativeInteger(JsonElement schema, string name) {
        if (schema.TryGetProperty(name, out JsonElement value) && (!value.TryGetInt32(out int parsed) || parsed < 0))
            throw new ArgumentException($"Tool-schema {name} must be a non-negative integer.", nameof(schema));
    }

    private static void ValidateNumber(JsonElement schema, string name) {
        if (schema.TryGetProperty(name, out JsonElement value) && value.ValueKind != JsonValueKind.Number)
            throw new ArgumentException($"Tool-schema {name} must be numeric.", nameof(schema));
    }

    private static InvalidOperationException Invalid(string path, string reason) =>
        new($"Tool arguments at '{path}' {reason}.");
}
