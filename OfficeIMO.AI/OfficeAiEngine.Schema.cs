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
        JsonNode rows = properties["tables"]!["items"]!["properties"]!["rows"]!;
        rows["maxItems"] = request.Limits.MaxResultItems;
        if (request.Operation == OfficeAiOperation.Parse && request.Limits.MaxTableCells < 100 * request.Limits.MaxResultItems) {
            // Any valid rectangle must fit at least one row-count/width pair. Group equivalent
            // row limits so narrow, tall tables remain possible without permitting oversized tables.
            var shapes = new JsonArray();
            int previousRows = 0;
            for (int width = 100; width >= 1; width--) {
                int count = Math.Min(request.Limits.MaxResultItems, request.Limits.MaxTableCells / width);
                if (count == previousRows) continue;
                previousRows = count;
                shapes.Add(new JsonObject {
                    ["type"] = "array", ["maxItems"] = count,
                    ["items"] = new JsonObject { ["type"] = "array", ["maxItems"] = width,
                        ["items"] = new JsonObject { ["type"] = "string" } }
                });
            }
            rows["anyOf"] = shapes;
        }
        return schema.ToJsonString();
    }

    private static string CreateSynthesisSchema(OfficeAiLimits limits) {
        JsonNode schema = JsonNode.Parse(SynthesisSchema)!;
        schema["properties"]!["claims"]!["maxItems"] = limits.MaxResultItems;
        return schema.ToJsonString();
    }
}
