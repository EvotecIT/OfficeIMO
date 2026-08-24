using System.Text.Json;
using System.Text.Json.Nodes;

namespace OfficeIMO.Adf.Benchmarks;

internal static class AdfComparisonValidation {
    internal static AdfOutputEvidence ValidateOfficeParse(string json, AdfDocument document) {
        AdfValidationResult result = document.Validate();
        if (!result.IsValid) throw new InvalidDataException("OfficeIMO parsed an invalid ADF document.");
        return Inspect(json, document.ToJson(), "OfficeIMO");
    }

    internal static AdfOutputEvidence ValidatePlatformParse(string json, JsonNode document) =>
        Inspect(json, document.ToJsonString(), "System.Text.Json");

    internal static AdfOutputEvidence Inspect(string input, string output, string implementation) {
        JsonNode inputNode = JsonNode.Parse(input) ?? throw new InvalidDataException("Input JSON is null.");
        JsonNode outputNode = JsonNode.Parse(output) ?? throw new InvalidDataException("Output JSON is null.");
        if (!JsonNode.DeepEquals(inputNode, outputNode)) {
            throw new InvalidDataException($"{implementation} did not preserve the ADF JSON tree.");
        }
        JsonObject root = outputNode.AsObject();
        if (root["sourceExtension"]?["owner"]?.GetValue<string>() != "OfficeIMO benchmark") {
            throw new InvalidDataException($"{implementation} did not preserve the root extension.");
        }
        JsonArray content = root["content"]?.AsArray()
            ?? throw new InvalidDataException($"{implementation} produced no content array.");
        JsonObject future = content[^1]?.AsObject()
            ?? throw new InvalidDataException($"{implementation} produced no future node.");
        if (future["futurePayload"]?.GetValue<string>() != "preserved") {
            throw new InvalidDataException($"{implementation} did not preserve the unknown node payload.");
        }
        return new AdfOutputEvidence(
            implementation,
            System.Text.Encoding.UTF8.GetByteCount(input),
            System.Text.Encoding.UTF8.GetByteCount(output),
            CountNodes(content),
            CountTextCharacters(content));
    }

    private static long CountNodes(JsonArray nodes) {
        long count = 0;
        foreach (JsonNode? node in nodes) {
            if (node is not JsonObject item) continue;
            count++;
            if (item["content"] is JsonArray children) count += CountNodes(children);
        }
        return count;
    }

    private static long CountTextCharacters(JsonArray nodes) {
        long count = 0;
        foreach (JsonNode? node in nodes) {
            if (node is not JsonObject item) continue;
            if (item["text"] is JsonValue text && text.TryGetValue(out string? value)) count += value?.Length ?? 0;
            if (item["content"] is JsonArray children) count += CountTextCharacters(children);
        }
        return count;
    }
}

internal sealed record AdfOutputEvidence(
    string Implementation,
    long InputBytes,
    long OutputBytes,
    long NodeCount,
    long TextCharacters);
