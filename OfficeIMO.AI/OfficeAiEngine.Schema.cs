using System.Text.Json.Nodes;

namespace OfficeIMO.AI;

public sealed partial class OfficeAiEngine {
    // The shared response envelope remains stable, but generation must obey the same
    // operation-specific empty-array contract that the local validator enforces.
    private static string CreateOutputSchema(OfficeAiRequest request) {
        string[] enabled = request.Operation switch {
            OfficeAiOperation.Ask or OfficeAiOperation.Explain or OfficeAiOperation.Summarize => new[] { "claims" },
            OfficeAiOperation.ExtractFields => new[] { "fields" },
            OfficeAiOperation.Parse => new[] { "blocks", "tables" },
            _ => throw new ArgumentOutOfRangeException(nameof(request))
        };
        JsonNode schema = JsonNode.Parse(Schema)!;
        JsonNode properties = schema["properties"]!;
        foreach (string name in new[] { "claims", "fields", "blocks", "tables" })
            properties[name]!["maxItems"] = enabled.Contains(name, StringComparer.Ordinal) ? request.Limits.MaxResultItems : 0;
        if (request.Operation == OfficeAiOperation.Parse) AddRectangularTableShapes(schema, request.Limits);
        BoundStrings(schema);
        return schema.ToJsonString();
    }

    private static void AddRectangularTableShapes(JsonNode schema, OfficeAiLimits limits) {
        JsonNode tables = schema["properties"]!["tables"]!;
        var definitions = new JsonObject {
            ["text"] = new JsonObject { ["type"] = "string" },
            ["evidence"] = tables["items"]!["properties"]!["evidence"]!.DeepClone()
        };
        var shapes = new JsonArray();
        for (int width = 1; width <= limits.MaxTableColumns; width++) {
            string rowDefinition = "row" + width;
            definitions[rowDefinition] = new JsonObject { ["type"] = "array", ["minItems"] = width, ["maxItems"] = width,
                ["items"] = Reference("text") };
            shapes.Add(new JsonObject {
                ["type"] = "object", ["additionalProperties"] = false,
                ["properties"] = new JsonObject {
                    ["title"] = Reference("text"), ["columns"] = Reference(rowDefinition),
                    ["rows"] = new JsonObject { ["type"] = "array",
                        ["maxItems"] = Math.Min(limits.MaxResultItems, limits.MaxTableCells / width),
                        ["items"] = Reference(rowDefinition) },
                    ["evidence"] = Reference("evidence")
                },
                ["required"] = new JsonArray("title", "columns", "rows", "evidence")
            });
        }
        schema["$defs"] = definitions;
        tables["items"] = new JsonObject { ["anyOf"] = shapes };
        static JsonObject Reference(string name) => new() { ["$ref"] = "#/$defs/" + name };
    }

    private static string CreateSynthesisSchema(OfficeAiLimits limits) {
        JsonNode schema = JsonNode.Parse(SynthesisSchema)!;
        schema["properties"]!["claims"]!["maxItems"] = limits.MaxResultItems;
        BoundStrings(schema);
        return schema.ToJsonString();
    }

    private const int MaxOutputStringLength = 32_000;
    // Match String.IsNullOrWhiteSpace's Unicode whitespace set independently of a provider's regex dialect.
    private const string NonblankStringPattern = @"^[^\u0000]*[^\u0000\u0009-\u000D\u0020\u0085\u00A0\u1680\u2000-\u200A\u2028\u2029\u202F\u205F\u3000][^\u0000]*$";

    private static void BoundStrings(JsonNode node, bool allowEmpty = false) {
        if (node is JsonObject obj) {
            if (obj["type"] is JsonValue scalar && scalar.TryGetValue<string>(out var type) && type == "string"
                || obj["type"] is JsonArray types && types.Any(value => value?.ToString() == "string")) {
                obj["maxLength"] = MaxOutputStringLength;
                obj["pattern"] = allowEmpty ? @"^[^\u0000]*$" : NonblankStringPattern;
                if (!allowEmpty) {
                    obj["minLength"] = 1;
                }
            }
            foreach (var property in obj) {
                if (property.Value is not { } value) continue;
                if (property.Key == "properties" && value is JsonObject properties) {
                    foreach (var field in properties)
                        if (field.Value is not null) BoundStrings(field.Value, field.Key is "title" or "columns" or "rows");
                } else if (property.Key == "$defs" && value is JsonObject definitions) {
                    foreach (var definition in definitions)
                        if (definition.Value is not null) BoundStrings(definition.Value, definition.Key == "text");
                } else BoundStrings(value, allowEmpty);
            }
        } else if (node is JsonArray array) {
            foreach (JsonNode? child in array) if (child is not null) BoundStrings(child, allowEmpty);
        }
    }
}
