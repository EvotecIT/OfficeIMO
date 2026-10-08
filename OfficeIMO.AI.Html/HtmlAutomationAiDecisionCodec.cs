using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using OfficeIMO.Html.Runtime;

namespace OfficeIMO.AI.Html;

// The bridge owns model response bounds; argument meaning stays with the HTML tool catalog.
internal static class HtmlAutomationAiDecisionCodec {
    internal static string Schema(IReadOnlyList<HtmlAutomationToolDefinition> tools, int maximumCalls) {
        var choices = new JsonArray();
        foreach (HtmlAutomationToolDefinition tool in tools) choices.Add(new JsonObject {
            ["type"] = "object", ["additionalProperties"] = false,
            ["required"] = new JsonArray("id", "name", "arguments"),
            ["properties"] = new JsonObject {
                ["id"] = new JsonObject { ["type"] = "string", ["minLength"] = 1 },
                ["name"] = new JsonObject { ["type"] = "string", ["enum"] = new JsonArray(tool.Name) },
                ["arguments"] = JsonNode.Parse(tool.InputSchema.GetRawText())
            }
        });
        return new JsonObject {
            ["type"] = "object", ["additionalProperties"] = false,
            ["required"] = new JsonArray("isComplete", "message", "calls"),
            ["properties"] = new JsonObject {
                ["isComplete"] = new JsonObject { ["type"] = "boolean" },
                ["message"] = new JsonObject { ["type"] = new JsonArray("string", "null") },
                ["calls"] = new JsonObject { ["type"] = "array", ["maxItems"] = maximumCalls,
                    ["items"] = new JsonObject { ["anyOf"] = choices } }
            }
        }.ToJsonString();
    }

    internal static HtmlAutomationPlannerDecision Parse(OfficeAiExecutionResponse response, HtmlAutomationAiPlannerOptions limits) {
        if (response == null || !response.IsComplete) throw Failure("The model response is missing or truncated.");
        if (response.Json == null || response.Json.Length > limits.MaxResponseCharacters)
            throw Failure("The model response exceeds the character limit.");
        JsonDocument document;
        try { document = JsonDocument.Parse(response.Json, new JsonDocumentOptions { MaxDepth = limits.MaxJsonDepth }); }
        catch (JsonException error) { throw new InvalidOperationException("The model decision is malformed or exceeds the JSON depth limit.", error); }
        using (document) {
            JsonElement root = document.RootElement;
            Object(root, "isComplete", "message", "calls");
            JsonElement status = Required(root, "isComplete");
            if (status.ValueKind is not (JsonValueKind.True or JsonValueKind.False)) throw Failure("The completion decision must be boolean.");
            JsonElement calls = Required(root, "calls");
            if (calls.ValueKind != JsonValueKind.Array || calls.GetArrayLength() > limits.MaxToolCallsPerTurn)
                throw Failure("The model decision exceeds the tool call limit.");
            string? message = null;
            JsonElement text = Required(root, "message");
            if (text.ValueKind != JsonValueKind.Null) {
                if (text.ValueKind != JsonValueKind.String) throw Failure("The decision message must be text or null.");
                message = text.GetString();
            }
            bool complete = status.GetBoolean();
            if (complete && calls.GetArrayLength() != 0 || !complete && calls.GetArrayLength() == 0)
                throw Failure("Completion decisions cannot contain calls, and continuing decisions require calls.");
            if (complete) return HtmlAutomationPlannerDecision.Complete(message);
            var ids = new HashSet<string>(StringComparer.Ordinal);
            var parsed = new List<HtmlAutomationToolCall>();
            foreach (JsonElement call in calls.EnumerateArray()) {
                Object(call, "id", "name", "arguments");
                string id = Text(call, "id"), name = Text(call, "name");
                if (!ids.Add(id)) throw Failure("The decision contains duplicate call identifiers.");
                if (!HtmlAutomationToolCatalog.GetDefinitions().Any(tool => tool.Name == name)) throw Failure("The decision names an undeclared tool.");
                JsonElement arguments = Required(call, "arguments");
                if (Encoding.UTF8.GetByteCount(arguments.GetRawText()) > limits.MaxArgumentBytes)
                    throw Failure("The tool arguments exceed the byte limit.");
                int count = 0;
                Count(arguments, limits.MaxArgumentItems, ref count);
                try { parsed.Add(new HtmlAutomationToolCall(id, name, HtmlAutomationToolCatalog.NormalizeArguments(name, arguments))); }
                catch (ArgumentException error) { throw new InvalidOperationException("The tool arguments are invalid: " + error.Message, error); }
            }
            return HtmlAutomationPlannerDecision.Execute(parsed.ToArray());
        }
    }

    private static void Object(JsonElement value, params string[] allowed) {
        if (value.ValueKind != JsonValueKind.Object) throw Failure("The decision requires JSON objects.");
        var names = new HashSet<string>(StringComparer.Ordinal);
        foreach (JsonProperty property in value.EnumerateObject()) {
            if (!names.Add(property.Name)) throw Failure("The decision contains a duplicate property.");
            if (!allowed.Contains(property.Name, StringComparer.Ordinal)) throw Failure("The decision contains an unknown property.");
        }
    }

    private static JsonElement Required(JsonElement value, string name) => value.TryGetProperty(name, out JsonElement result)
        ? result : throw Failure("The decision is missing a required property.");

    private static string Text(JsonElement value, string name) {
        JsonElement property = Required(value, name);
        if (property.ValueKind != JsonValueKind.String || string.IsNullOrWhiteSpace(property.GetString()))
            throw Failure("Call identifiers and names must be nonempty text.");
        return property.GetString()!;
    }

    private static void Count(JsonElement value, int maximum, ref int count) {
        if (value.ValueKind == JsonValueKind.Object) {
            foreach (JsonProperty property in value.EnumerateObject()) {
                if (++count > maximum) throw Failure("The tool arguments exceed the item limit.");
                Count(property.Value, maximum, ref count);
            }
        } else if (value.ValueKind == JsonValueKind.Array) {
            foreach (JsonElement item in value.EnumerateArray()) {
                if (++count > maximum) throw Failure("The tool arguments exceed the item limit.");
                Count(item, maximum, ref count);
            }
        }
    }

    private static InvalidOperationException Failure(string message) => new(message);
}
