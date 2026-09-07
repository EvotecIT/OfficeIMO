using System.Text.Json.Nodes;

namespace OfficeIMO.AI;

public sealed partial class OfficeAiEngine {
    // The shared response envelope remains stable, but generation must obey the same
    // operation-specific empty-array contract that the local validator enforces.
    private static string CreateOutputSchema(OfficeAiOperation operation) {
        string[] enabled = operation switch {
            OfficeAiOperation.Ask or OfficeAiOperation.Explain or OfficeAiOperation.Summarize => new[] { "claims" },
            OfficeAiOperation.ExtractFields => new[] { "fields" },
            OfficeAiOperation.Parse => new[] { "blocks", "tables" },
            _ => throw new ArgumentOutOfRangeException(nameof(operation))
        };
        JsonNode schema = JsonNode.Parse(Schema)!;
        JsonNode properties = schema["properties"]!;
        foreach (string name in new[] { "claims", "fields", "blocks", "tables" })
            if (!enabled.Contains(name, StringComparer.Ordinal)) properties[name]!["maxItems"] = 0;
        return schema.ToJsonString();
    }
}
